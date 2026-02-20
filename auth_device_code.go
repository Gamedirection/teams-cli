package main

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/rivo/tview"
)

const (
	teamsClientID = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346"

	skypeResource    = "https://api.spaces.skype.com"
	chatSvcResource  = "https://chatsvcagg.teams.microsoft.com"
	teamsTokenPrefix = "token-"
)

type deviceCodeResponse struct {
	DeviceCode              string `json:"device_code"`
	UserCode                string `json:"user_code"`
	VerificationURI         string `json:"verification_uri"`
	VerificationURIComplete string `json:"verification_uri_complete"`
	ExpiresIn               int    `json:"expires_in"`
	Interval                int    `json:"interval"`
	Message                 string `json:"message"`
	Error                   string `json:"error"`
	ErrorDescription        string `json:"error_description"`
}

type tokenResponse struct {
	TokenType        string `json:"token_type"`
	ExpiresIn        int    `json:"expires_in"`
	AccessToken      string `json:"access_token"`
	RefreshToken     string `json:"refresh_token"`
	IDToken          string `json:"id_token"`
	Scope            string `json:"scope"`
	Error            string `json:"error"`
	ErrorDescription string `json:"error_description"`
}

func (s *AppState) refreshAuthFromDeviceCode() error {
	if disableDeviceCode() {
		return fmt.Errorf("device code auth disabled")
	}

	tenant := strings.TrimSpace(os.Getenv("TEAMS_CLI_TENANT"))
	if tenant == "" {
		tenant = "common"
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Minute)
	defer cancel()

	codeResp, err := requestDeviceCode(ctx, tenant, "openid profile offline_access")
	if err != nil {
		return err
	}

	s.renderDeviceCodeMessage(codeResp)

	tokenResp, err := pollDeviceCode(ctx, tenant, codeResp)
	if err != nil {
		return err
	}

	teamsToken := strings.TrimSpace(tokenResp.IDToken)
	if teamsToken == "" {
		teamsToken = strings.TrimSpace(tokenResp.AccessToken)
	}
	if teamsToken == "" {
		return fmt.Errorf("device code flow did not return an id_token or access_token")
	}

	if err := writeTokenFile("teams", teamsToken); err != nil {
		return err
	}

	if tokenResp.RefreshToken == "" {
		return fmt.Errorf("device code flow did not return a refresh_token")
	}

	if err := refreshResourceToken(ctx, tenant, tokenResp.RefreshToken, skypeResource, "skype"); err != nil {
		return err
	}
	if err := refreshResourceToken(ctx, tenant, tokenResp.RefreshToken, chatSvcResource, "chatsvcagg"); err != nil {
		return err
	}

	return nil
}

func disableDeviceCode() bool {
	val := strings.ToLower(strings.TrimSpace(os.Getenv("TEAMS_CLI_DISABLE_DEVICE_CODE")))
	return val == "1" || val == "true" || val == "yes"
}

func (s *AppState) renderDeviceCodeMessage(resp *deviceCodeResponse) {
	if resp == nil || s == nil {
		return
	}
	msg := strings.TrimSpace(resp.Message)
	if msg == "" {
		if resp.VerificationURIComplete != "" {
			msg = fmt.Sprintf("Go to %s", resp.VerificationURIComplete)
		} else if resp.VerificationURI != "" && resp.UserCode != "" {
			msg = fmt.Sprintf("Go to %s and enter code %s", resp.VerificationURI, resp.UserCode)
		}
	}
	if msg == "" {
		return
	}

	if val, ok := s.components[TvError]; ok {
		if tv, ok := val.(*tview.TextView); ok {
			s.app.QueueUpdateDraw(func() {
				tv.SetText(msg)
			})
		}
	}
}

func requestDeviceCode(ctx context.Context, tenant string, scope string) (*deviceCodeResponse, error) {
	form := url.Values{}
	form.Set("client_id", teamsClientID)
	form.Set("scope", scope)

	u := fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/devicecode", url.PathEscape(tenant))
	resp, err := httpPostForm(ctx, u, form)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	var payload deviceCodeResponse
	if err := json.Unmarshal(body, &payload); err != nil {
		return nil, fmt.Errorf("device code response parse failed: %v", err)
	}
	if payload.Error != "" {
		return nil, fmt.Errorf("device code error: %s (%s)", payload.Error, payload.ErrorDescription)
	}
	if payload.DeviceCode == "" {
		return nil, errors.New("device code response missing device_code")
	}
	if payload.Interval <= 0 {
		payload.Interval = 5
	}
	return &payload, nil
}

func pollDeviceCode(ctx context.Context, tenant string, device *deviceCodeResponse) (*tokenResponse, error) {
	if device == nil {
		return nil, errors.New("device code is nil")
	}

	u := fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/token", url.PathEscape(tenant))
	deadline := time.Now().Add(time.Duration(device.ExpiresIn) * time.Second)
	interval := time.Duration(device.Interval) * time.Second
	if interval < time.Second {
		interval = time.Second
	}

	for {
		if time.Now().After(deadline) {
			return nil, errors.New("device code expired before authorization completed")
		}
		if ctx.Err() != nil {
			return nil, ctx.Err()
		}

		form := url.Values{}
		form.Set("grant_type", "urn:ietf:params:oauth:grant-type:device_code")
		form.Set("client_id", teamsClientID)
		form.Set("device_code", device.DeviceCode)

		resp, err := httpPostForm(ctx, u, form)
		if err != nil {
			return nil, err
		}
		body, _ := io.ReadAll(resp.Body)
		resp.Body.Close()

		var payload tokenResponse
		if err := json.Unmarshal(body, &payload); err != nil {
			return nil, fmt.Errorf("token response parse failed: %v", err)
		}

		if payload.Error == "" && payload.AccessToken != "" {
			return &payload, nil
		}

		switch payload.Error {
		case "authorization_pending":
			time.Sleep(interval)
			continue
		case "slow_down":
			interval += time.Second
			time.Sleep(interval)
			continue
		case "authorization_declined":
			return nil, errors.New("device code authorization declined")
		case "expired_token":
			return nil, errors.New("device code expired")
		case "":
			return nil, errors.New("device code token response missing access_token")
		default:
			return nil, fmt.Errorf("device code token error: %s (%s)", payload.Error, payload.ErrorDescription)
		}
	}
}

func refreshResourceToken(ctx context.Context, tenant, refreshToken, resource, label string) error {
	scope := fmt.Sprintf("%s/.default", resource)
	form := url.Values{}
	form.Set("grant_type", "refresh_token")
	form.Set("client_id", teamsClientID)
	form.Set("refresh_token", refreshToken)
	form.Set("scope", scope)

	u := fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/v2.0/token", url.PathEscape(tenant))
	resp, err := httpPostForm(ctx, u, form)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	var payload tokenResponse
	if err := json.Unmarshal(body, &payload); err != nil {
		return fmt.Errorf("refresh token response parse failed: %v", err)
	}
	if payload.Error != "" {
		return fmt.Errorf("refresh token error for %s: %s (%s)", label, payload.Error, payload.ErrorDescription)
	}
	if payload.AccessToken == "" {
		return fmt.Errorf("refresh token response missing access_token for %s", label)
	}
	return writeTokenFile(label, payload.AccessToken)
}

func writeTokenFile(kind, token string) error {
	home, err := os.UserHomeDir()
	if err != nil {
		return err
	}
	dir := filepath.Join(home, ".config", "fossteams")
	if err := os.MkdirAll(dir, 0o700); err != nil {
		return err
	}
	path := filepath.Join(dir, teamsTokenPrefix+kind+".jwt")
	return os.WriteFile(path, []byte(token), 0o600)
}

func httpPostForm(ctx context.Context, endpoint string, form url.Values) (*http.Response, error) {
	req, err := http.NewRequestWithContext(ctx, http.MethodPost, endpoint, strings.NewReader(form.Encode()))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	client := &http.Client{Timeout: 30 * time.Second}
	return client.Do(req)
}
