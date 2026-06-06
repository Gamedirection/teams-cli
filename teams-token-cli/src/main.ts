import puppeteer, { Browser, Page, Frame } from 'puppeteer-core';
import { homedir } from 'os';
import { writeFileSync, existsSync, mkdirSync, unlinkSync } from 'fs';
import { join } from 'path';
import { v4 as uuidv4 } from 'uuid';
import axios from 'axios';
import * as jwt from 'jsonwebtoken';

// eslint-disable-next-line @typescript-eslint/no-var-requires
const carbonylBin: string = require('carbonyl');

const CONFIG_DIR = join(homedir(), '.config', 'fossteams');
const MICROSOFT_TENANT_ID = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a';
const TEAMS_APP_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const SKYPE_RESOURCE = 'https://api.spaces.skype.com';
const CHAT_SVC_AGG_RESOURCE = 'https://chatsvcagg.teams.microsoft.com';
const USER_AGENT = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) MicrosoftTeams-Preview/1.4.00.7556 Chrome/80.0.3987.163 Electron/8.5.5 Safari/537.36';

type TeamsSkype = 'teams' | 'skype' | 'chatsvcagg';

function getLoginURL(type: TeamsSkype, tenantId: string): string {
  const loginUrl = new URL('https://login.microsoftonline.com');
  loginUrl.pathname = `/${tenantId}/oauth2/authorize`;

  const state = uuidv4();
  switch (type) {
    case 'teams':
      loginUrl.searchParams.append('response_type', 'id_token');
      loginUrl.searchParams.append('state', state);
      break;
    case 'skype':
      loginUrl.searchParams.append('response_type', 'token');
      loginUrl.searchParams.append('state', `${state}|${SKYPE_RESOURCE}`);
      loginUrl.searchParams.append('resource', SKYPE_RESOURCE);
      break;
    case 'chatsvcagg':
      loginUrl.searchParams.append('response_type', 'token');
      loginUrl.searchParams.append('state', `${state}|${CHAT_SVC_AGG_RESOURCE}`);
      loginUrl.searchParams.append('resource', CHAT_SVC_AGG_RESOURCE);
      break;
    default:
      break;
  }
  loginUrl.searchParams.append('client_id', TEAMS_APP_ID);
  loginUrl.searchParams.append('client-request-id', uuidv4());
  loginUrl.searchParams.append('redirect_uri', 'https://teams.microsoft.com/go');
  loginUrl.searchParams.append('x-client-SKU', 'Js');
  loginUrl.searchParams.append('x-client-Ver', '1.0.9');
  loginUrl.searchParams.append('nonce', uuidv4());

  return loginUrl.toString();
}

function saveToken(token: string, type: TeamsSkype): void {
  if (!existsSync(CONFIG_DIR)) {
    mkdirSync(CONFIG_DIR, { recursive: true });
  }
  writeFileSync(join(CONFIG_DIR, `token-${type}.jwt`), token);
}

function getTenants(token: string) {
  return axios.get<Array<{ tenantId: string }>>('https://teams.microsoft.com/api/mt/emea/beta/users/tenants', {
    headers: { Authorization: `Bearer ${token}` },
  });
}

// Waits for navigation to teams.microsoft.com/go and extracts the token from the URL hash.
function waitForRedirect(page: Page): Promise<string> {
  return new Promise((resolve, reject) => {
    let settled = false;

    const handler = (frame: Frame) => {
      if (frame !== page.mainFrame()) return;
      const url = frame.url();
      if (!url.startsWith('https://teams.microsoft.com/go')) return;

      page.off('framenavigated', handler);
      if (settled) return;
      settled = true;

      const hash = url.includes('#') ? url.split('#')[1] : '';
      if (!hash) {
        // Try to get hash from page JS in case framenavigated fired before hash was set
        page.evaluate(() => window.location.hash.replace(/^#/, '')).then(h => {
          resolve(h || '');
        }).catch(() => resolve(''));
        return;
      }
      resolve(hash);
    };

    page.on('framenavigated', handler);

    // Timeout after 5 minutes
    setTimeout(() => {
      if (!settled) {
        settled = true;
        page.off('framenavigated', handler);
        reject(new Error('Timed out waiting for Teams auth redirect'));
      }
    }, 5 * 60 * 1000);
  });
}

async function authorize(page: Page, type: TeamsSkype, tenantId: string): Promise<string> {
  console.log(`Authorizing ${type} with tenantId=${tenantId}`);
  const redirectPromise = waitForRedirect(page);
  await page.setUserAgent(USER_AGENT);
  // Don't await goto — it may never resolve because we intercept the redirect
  page.goto(getLoginURL(type, tenantId)).catch(() => {});
  return redirectPromise;
}

async function main() {
  const browser: Browser = await puppeteer.launch({
    executablePath: carbonylBin,
    headless: false,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu-sandbox'],
  });

  try {
    const page = await browser.newPage();
    await page.setUserAgent(USER_AGENT);

    let currentTenant: string | null = null;
    let redirectCount = 0;

    const processHash = async (hash: string, type: TeamsSkype): Promise<void> => {
      redirectCount++;
      if (redirectCount > 8) throw new Error('Too many redirects');

      const searchParams = new URLSearchParams(hash);
      const token = searchParams.get('id_token') ?? searchParams.get('access_token');

      if (!token) {
        if (searchParams.has('error')) {
          throw new Error(`${searchParams.get('error')}: ${searchParams.get('error_description')}`);
        }
        throw new Error('No token in redirect URL');
      }

      const decoded = jwt.decode(token) as Record<string, unknown>;
      if (!decoded || typeof decoded === 'string') throw new Error('Invalid JWT');

      console.log(`Audience: ${decoded.aud}`);
      console.log('Decoded', decoded);

      // Microsoft-tenant Skype response → need to resolve real tenant first
      if (decoded.tid === MICROSOFT_TENANT_ID && decoded.aud === SKYPE_RESOURCE) {
        console.log('Getting tenant list...');
        const tenants = (await getTenants(token)).data;
        currentTenant = tenants[0].tenantId;
        const skypeHash = await authorize(page, 'skype', currentTenant);
        return processHash(skypeHash, 'skype');
      }

      if (!currentTenant) {
        currentTenant = (decoded.tid !== MICROSOFT_TENANT_ID ? decoded.tid : 'common') as string;
      }

      if (decoded.aud === TEAMS_APP_ID) {
        console.log('Got Teams token');
        saveToken(token, 'teams');
        const skypeHash = await authorize(page, 'skype', currentTenant);
        return processHash(skypeHash, 'skype');
      } else if (decoded.aud === SKYPE_RESOURCE) {
        console.log('Got Skype token');
        saveToken(token, 'skype');
        const aggHash = await authorize(page, 'chatsvcagg', currentTenant);
        return processHash(aggHash, 'chatsvcagg');
      } else if (decoded.aud === CHAT_SVC_AGG_RESOURCE) {
        console.log('Got ChatSvcAgg token');
        saveToken(token, 'chatsvcagg');
      } else {
        throw new Error(`Unknown audience: ${decoded.aud}`);
      }
    };

    const teamsHash = await authorize(page, 'teams', 'common');
    await processHash(teamsHash, 'teams');

    console.log('\nAll tokens saved to', CONFIG_DIR);
  } finally {
    await browser.close();
  }
}

// CLI entry
const arg = process.argv[2];

if (arg === 'logout') {
  ['teams', 'skype', 'chatsvcagg'].forEach(type => {
    const f = join(CONFIG_DIR, `token-${type}.jwt`);
    if (existsSync(f)) {
      unlinkSync(f);
      console.log(`Removed ${f}`);
    }
  });
  process.exit(0);
} else if (arg === 'get-url') {
  console.log(getLoginURL('teams', 'common'));
  process.exit(0);
} else {
  main().catch(err => {
    console.error('Error:', err.message || err);
    process.exit(1);
  });
}
