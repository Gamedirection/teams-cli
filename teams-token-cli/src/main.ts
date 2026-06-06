import puppeteer, { Browser, Page, Frame } from 'puppeteer-core';
import { spawn } from 'child_process';
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
const DEBUG_PORT = 9977;

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

    const settle = (hash: string) => {
      if (settled) return;
      settled = true;
      // Navigate away immediately to prevent Teams web app loading in terminal
      page.goto('about:blank').catch(() => {});
      resolve(hash);
    };

    // Primary: intercept the 302 Location header — hash is guaranteed to be here
    const responseHandler = async (response) => {
      const status = response.status();
      if (![301, 302, 303, 307, 308].includes(status)) return;
      const location = response.headers()['location'] || '';
      if (!location.startsWith('https://teams.microsoft.com/go')) return;
      page.off('response', responseHandler);
      page.off('framenavigated', navHandler);
      const hash = location.includes('#') ? location.split('#')[1] : '';
      settle(hash);
    };

    // Fallback: framenavigated in case response event misses it
    const navHandler = async (frame: Frame) => {
      if (frame !== page.mainFrame()) return;
      const url = frame.url();
      if (!url.startsWith('https://teams.microsoft.com/go')) return;
      page.off('response', responseHandler);
      page.off('framenavigated', navHandler);
      const hash = url.includes('#') ? url.split('#')[1] : '';
      settle(hash);
    };

    page.on('response', responseHandler);
    page.on('framenavigated', navHandler);

    setTimeout(() => {
      if (!settled) {
        settled = true;
        page.off('response', responseHandler);
        page.off('framenavigated', navHandler);
        reject(new Error('Timed out waiting for Teams auth redirect (5 min)'));
      }
    }, 5 * 60 * 1000);
  });
}

async function authorize(page: Page, type: TeamsSkype, tenantId: string): Promise<string> {
  log(`Authorizing ${type} with tenantId=${tenantId}`);
  const redirectPromise = waitForRedirect(page);
  await page.setUserAgent(USER_AGENT);
  // timeout:0 — user may take minutes to log in; default 30s kills the page
  page.goto(getLoginURL(type, tenantId), { timeout: 0 }).catch(() => {});
  return redirectPromise;
}

async function waitForCDP(port: number): Promise<void> {
  for (let i = 0; i < 75; i++) {
    try {
      const res = await fetch(`http://localhost:${port}/json`);
      if (res.ok) return;
    } catch { /* not ready yet */ }
    await new Promise(r => setTimeout(r, 200));
  }
  throw new Error(`Carbonyl CDP not available on port ${port} after 15s`);
}

const log = (msg: string) => process.stderr.write(msg + '\n');

async function main() {
  const initialURL = getLoginURL('teams', 'common');

  // Pass initial URL so Carbonyl starts rendering immediately instead of blank
  const proc = spawn(carbonylBin, [
    `--remote-debugging-port=${DEBUG_PORT}`,
    '--no-sandbox',
    '--disable-setuid-sandbox',
    initialURL,
  ], { stdio: 'inherit' });

  proc.on('error', (err) => { log('Failed to start Carbonyl: ' + err.message); process.exit(1); });

  const cleanup = () => { try { proc.kill(); } catch { /* ignore */ } };
  process.on('exit', cleanup);
  process.on('SIGINT', () => { cleanup(); process.exit(0); });
  process.on('SIGTERM', () => { cleanup(); process.exit(0); });

  await waitForCDP(DEBUG_PORT);

  const browser: Browser = await puppeteer.connect({
    browserURL: `http://localhost:${DEBUG_PORT}`,
  });

  try {
    const pages = await browser.pages();
    const page = pages[0] || await browser.newPage();
    await page.setUserAgent(USER_AGENT);

    let currentTenant: string | null = null;
    let redirectCount = 0;

    const processHash = async (hash: string): Promise<void> => {
      redirectCount++;
      if (redirectCount > 8) throw new Error('Too many redirects');

      const searchParams = new URLSearchParams(hash);
      const token = searchParams.get('id_token') ?? searchParams.get('access_token');

      if (!token) {
        if (searchParams.has('error')) {
          throw new Error(`${searchParams.get('error')}: ${searchParams.get('error_description')}`);
        }
        throw new Error('No token found in redirect URL');
      }

      const decoded = jwt.decode(token) as Record<string, unknown>;
      if (!decoded || typeof decoded === 'string') throw new Error('Invalid JWT');

      log(`Audience: ${decoded.aud}`);

      if (decoded.tid === MICROSOFT_TENANT_ID && decoded.aud === SKYPE_RESOURCE) {
        log('Resolving real tenant...');
        const tenants = (await getTenants(token)).data;
        currentTenant = tenants[0].tenantId;
        return processHash(await authorize(page, 'skype', currentTenant));
      }

      if (!currentTenant) {
        currentTenant = (decoded.tid !== MICROSOFT_TENANT_ID ? decoded.tid : 'common') as string;
      }

      if (decoded.aud === TEAMS_APP_ID) {
        log('Got Teams token');
        saveToken(token, 'teams');
        return processHash(await authorize(page, 'skype', currentTenant));
      } else if (decoded.aud === SKYPE_RESOURCE) {
        log('Got Skype token');
        saveToken(token, 'skype');
        return processHash(await authorize(page, 'chatsvcagg', currentTenant));
      } else if (decoded.aud === CHAT_SVC_AGG_RESOURCE) {
        log('Got ChatSvcAgg token');
        saveToken(token, 'chatsvcagg');
      } else {
        throw new Error(`Unknown audience: ${decoded.aud}`);
      }
    };

    await processHash(await authorize(page, 'teams', 'common'));
    log(`\nAll tokens saved to ${CONFIG_DIR}`);
  } finally {
    browser.disconnect();
    cleanup();
  }
}

// CLI entry
const arg = process.argv[2];

if (arg === 'logout') {
  ['teams', 'skype', 'chatsvcagg'].forEach(type => {
    const f = join(CONFIG_DIR, `token-${type}.jwt`);
    if (existsSync(f)) {
      unlinkSync(f);
      log(`Removed ${f}`);
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
