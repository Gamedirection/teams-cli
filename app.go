package main

import (
	"bytes"
	"crypto/aes"
	"crypto/cipher"
	"crypto/rand"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"github.com/dgrijalva/jwt-go"
	teams_api "github.com/fossteams/teams-api"
	api "github.com/fossteams/teams-api/pkg"
	"github.com/fossteams/teams-api/pkg/csa"
	"github.com/fossteams/teams-api/pkg/models"
	"github.com/gdamore/tcell/v2"
	"github.com/rivo/tview"
	"github.com/sirupsen/logrus"
	"golang.org/x/net/html"
	"io"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"runtime/debug"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

type AppState struct {
	app    *tview.Application
	pages  *tview.Pages
	logger *logrus.Logger

	TeamsState
	components map[string]tview.Primitive

	activeConversationMu    sync.RWMutex
	activeConversationIDs   []string
	activeConversationTitle string
	activeConversationNode  *tview.TreeNode

	chatFavoritesMu sync.RWMutex
	chatFavorites   map[string]bool
	chatTitlesMu    sync.RWMutex
	chatTitles      map[string]string

	authRefreshMu sync.Mutex

	settingsPath string
	settingsKey  string
	keybindPath  string

	unreadScanMu       sync.RWMutex
	unreadScanEnabled  bool
	unreadScanInterval time.Duration
	unreadScanStop     chan struct{}
	unreadScanRunning  bool
	unreadLastScanAt   time.Time
	unreadLastChanges  int

	manualUnreadMu sync.RWMutex
	manualUnread   map[string]bool

	chatMessagesMu sync.RWMutex
	chatMessages   []csa.ChatMessage
	chatRowMap     []int

	replyMu      sync.RWMutex
	pendingReply *replyTarget

	messageReactionsMu sync.RWMutex
	messageReactions   map[string]string

	keybindMu        sync.RWMutex
	keybindings      map[string][]string
	keybindPreset    string
	keybindOverrides map[string][]string
	keybindParseErr  error

	settingsMu            sync.RWMutex
	settingsMode          bool
	settingsSelection     int
	settingsCaptureAction string

	mentionCycleMu    sync.Mutex
	mentionCycleToken string
	mentionCycleIndex int
	mentionCycleItems []mentionCandidate

	contactsMu          sync.RWMutex
	contactCandidates   []mentionCandidate
	contactsLastFetched time.Time

	chatWordWrapMu    sync.RWMutex
	chatWordWrap      bool
	chatWrapChars     int
	chatWrapEffective int

	themeMu          sync.RWMutex
	composeColorName string
	authorColorName  string
}

type conversationRef struct {
	ids        []string
	title      string
	chatKey    string
	isFavorite bool
	isUnread   bool
}

type replyTarget struct {
	MessageID string
	Author    string
	Preview   string
}

type mentionCandidate struct {
	DisplayName string
	Mri         string
	ObjectID    string
}

type mentionWire struct {
	ID          int    `json:"id"`
	MentionType string `json:"mentionType"`
	Mri         string `json:"mri,omitempty"`
	DisplayName string `json:"displayName"`
	ObjectId    string `json:"objectId,omitempty"`
}

type encryptedSettingsFile struct {
	Version    int    `json:"version"`
	Nonce      string `json:"nonce"`
	Ciphertext string `json:"ciphertext"`
}

type persistedChatSettings struct {
	Favorites       map[string]bool   `json:"favorites"`
	Titles          map[string]string `json:"titles"`
	UnreadOverrides map[string]bool   `json:"unread_overrides,omitempty"`
	ChatWordWrap    *bool             `json:"chat_word_wrap,omitempty"`
	ChatWrapPercent *int              `json:"chat_wrap_percent,omitempty"`
	ChatWrapChars   *int              `json:"chat_wrap_chars,omitempty"`
	ComposeColor    string            `json:"compose_color,omitempty"`
	AuthorColor     string            `json:"author_color,omitempty"`
}

type keybindingConfigFile struct {
	Preset   string              `json:"preset"`
	Bindings map[string][]string `json:"bindings"`
}

type settingsItem struct {
	kind   string
	action string
}

const composeDefaultPlaceholder = "Press i to compose, f to toggle favorite chat, Enter to send, Esc to return"
const settingsHelpChatKey = "__settings_help__"
const defaultReactionKey = "like"
const defaultKeybindPreset = "default"

var mentionTokenRegex = regexp.MustCompile(`(?m)(^|[\s])([cC]?@)([A-Za-z0-9._-]+)`)

const (
	settingsItemInfo         = "info"
	settingsItemSpacer       = "spacer"
	settingsItemOpen         = "open_editor"
	settingsItemPreset       = "preset"
	settingsItemReload       = "reload"
	settingsItemBinding      = "binding"
	settingsItemWrap         = "chat_wrap"
	settingsItemWrapPct      = "chat_wrap_pct"
	settingsItemComposeColor = "compose_color"
	settingsItemAuthorColor  = "author_color"
)

const (
	actionToggleScan     = "toggle_scan"
	actionScanNow        = "scan_now"
	actionMarkUnread     = "mark_unread"
	actionToggleFavorite = "toggle_favorite"
	actionRefreshTitles  = "refresh_titles"
	actionReloadKeybinds = "reload_keybindings"
	actionFocusCompose   = "focus_compose"
	actionReplyMessage   = "reply_message"
	actionReactMessage   = "react_message"
	actionMoveDown       = "move_down"
	actionMoveUp         = "move_up"
)

func (s *AppState) createApp() {
	s.logger.Debug("creating application pages and components")
	s.pages = tview.NewPages()
	s.components = map[string]tview.Primitive{}
	s.chatFavorites = map[string]bool{}
	s.chatTitles = map[string]string{}
	s.unreadScanEnabled = true
	s.unreadScanInterval = time.Minute
	s.unreadScanStop = make(chan struct{})
	s.manualUnread = map[string]bool{}
	s.messageReactions = map[string]string{}
	s.chatWordWrap = true
	s.chatWrapChars = 80
	s.composeColorName = "slate"
	s.authorColorName = "blue"
	s.settingsPath, s.settingsKey = defaultSettingsPaths()
	s.keybindPath = defaultKeybindPath()
	s.keybindPreset = defaultKeybindPreset
	s.keybindings = defaultKeybindingsForPreset(defaultKeybindPreset)
	s.keybindOverrides = map[string][]string{}
	if err := s.loadKeybindingsConfig(); err != nil {
		s.keybindParseErr = err
		s.logger.WithError(err).Warn("unable to load keybinding config")
	}
	if err := s.loadEncryptedChatSettings(); err != nil {
		s.logger.WithError(err).Warn("unable to load encrypted chat settings")
	}

	// Add pages
	s.pages.AddPage(PageLogin, s.createLoginPage(), true, false)
	s.pages.AddPage(PageMain, s.createMainView(), true, false)
	s.pages.AddPage(PageError, s.createErrorView(), true, false)

	frame := tview.NewFrame(s.pages)
	frame.SetBorder(true)
	frame.SetTitle("teams-cli")
	frame.SetBorder(true)
	frame.SetTitleAlign(tview.AlignCenter)

	s.app.SetRoot(frame, true)

	// Set main page
	s.pages.SwitchToPage(PageLogin)
	s.app.SetFocus(s.pages)
	s.app.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		switch event.Key() {
		case tcell.KeyTAB:
			s.focusNextPane()
			return nil
		case tcell.KeyBacktab:
			s.focusPrevPane()
			return nil
		default:
			return event
		}
	})

	s.logger.Debug("starting async app initialization")
	go s.start()
}

func (s *AppState) focusNextPane() {
	tree := s.components[TrChat]
	chat := s.components[ViChat]
	compose := s.components[ViCompose]
	current := s.app.GetFocus()

	switch current {
	case tree:
		s.app.SetFocus(chat)
	case chat:
		s.app.SetFocus(compose)
	default:
		s.app.SetFocus(tree)
	}
}

func (s *AppState) focusPrevPane() {
	tree := s.components[TrChat]
	chat := s.components[ViChat]
	compose := s.components[ViCompose]
	current := s.app.GetFocus()

	switch current {
	case compose:
		s.app.SetFocus(chat)
	case chat:
		s.app.SetFocus(tree)
	default:
		s.app.SetFocus(compose)
	}
}

func (s *AppState) createMainView() tview.Primitive {
	// Top: User information
	// Left side: Tree view (Teams _ Channels / Conversations)
	// Right side: Chat view
	// Bottom: Navigation bar

	treeView := tview.NewTreeView()
	chatView := tview.NewList()
	chatView.SetBackgroundColor(tcell.ColorBlack)
	composeView := tview.NewInputField().
		SetLabel("Message: ").
		SetPlaceholder(composeDefaultPlaceholder)
	composeView.SetFieldWidth(0)
	composeView.SetFieldBackgroundColor(s.composeFieldColor())
	composeView.SetFieldTextColor(tcell.ColorWhite)
	composeView.SetLabelColor(tcell.ColorWhite)
	composeView.SetPlaceholderTextColor(tcell.ColorLightGray)
	composeView.SetBorder(true)
	composeView.SetTitle(s.composeTitleWithScanStatus())
	composeView.SetTitleAlign(tview.AlignCenter)

	s.components[TrChat] = treeView
	s.components[ViChat] = chatView
	s.components[ViCompose] = composeView

	chatPane := tview.NewFlex().SetDirection(tview.FlexRow).
		AddItem(chatView, 0, 1, false).
		AddItem(composeView, 3, 0, false)

	flex := tview.NewFlex().
		AddItem(treeView, 0, 1, false).
		AddItem(chatPane, 0, 2, false)

	return flex
}

func (s *AppState) createLoginPage() tview.Primitive {
	p := tview.NewTextView()
	p.SetTitle("Log-in")
	p.SetText("Logging in...")
	p.SetBackgroundColor(tcell.ColorBlue)
	p.SetTextAlign(tview.AlignCenter)
	p.SetBorder(true)
	p.SetBorderPadding(1, 1, 1, 1)

	s.components[TvLoginStatus] = p

	return tview.NewFlex().
		AddItem(nil, 0, 1, false).
		AddItem(tview.NewFlex().SetDirection(tview.FlexRow).
			AddItem(nil, 0, 1, false).
			AddItem(p, 10, 1, false).
			AddItem(nil, 0, 1, false), 30, 1, false).
		AddItem(nil, 0, 1, false)
}

func (s *AppState) start() {
	s.logTokenDiagnostics()

	s.logger.Info("initializing Teams client")
	// Initialize Teams client
	var err error
	s.teamsClient, err = teams_api.New()
	if err != nil {
		s.logger.WithError(err).Error("teams client initialization failed")
		s.showError(err)
		return
	}
	s.logger.Info("Teams client initialized")

	// Initialize Teams State
	s.TeamsState.logger = s.logger
	s.logger.Info("initializing Teams state")
	err = s.TeamsState.init(s.teamsClient)
	if err != nil {
		s.logger.WithError(err).Error("teams state initialization failed")
		s.showError(err)
		return
	}
	s.logger.Info("Teams state initialized")

	go s.fillMainWindow()
}

func (s *AppState) logTokenDiagnostics() {
	s.logger.Info("running token/auth diagnostics")

	skypeSpacesToken, err := api.GetSkypeSpacesToken()
	if err != nil {
		s.logger.WithError(err).Error("unable to load skype token")
	} else {
		logTokenMeta(s.logger, "skype", skypeSpacesToken)
	}

	chatSvcToken, err := api.GetChatSvcAggToken()
	if err != nil {
		s.logger.WithError(err).Error("unable to load chatsvcagg token")
	} else {
		logTokenMeta(s.logger, "chatsvcagg", chatSvcToken)
	}

	_, err = api.GetSkypeToken()
	if err != nil {
		s.logger.WithError(err).Error("unable to refresh skype token via authz")
	} else {
		s.logger.Info("skype token refresh via authz succeeded")
	}
}

func logTokenMeta(logger *logrus.Logger, name string, token *api.TeamsToken) {
	if logger == nil {
		return
	}
	if token == nil || token.Inner == nil {
		logger.WithField("token_name", name).Warn("token is nil")
		return
	}
	exp, ok := tokenExpiry(token)
	if !ok {
		logger.WithField("token_name", name).Warn("token has no parseable exp claim")
		return
	}
	logger.WithFields(logrus.Fields{
		"token_name":   name,
		"expires_at":   exp.Format(time.RFC3339),
		"is_expired":   time.Now().After(exp),
		"minutes_left": int(time.Until(exp).Minutes()),
	}).Info("token metadata")
}

func tokenExpiry(token *api.TeamsToken) (time.Time, bool) {
	if token == nil || token.Inner == nil {
		return time.Time{}, false
	}
	claims, ok := token.Inner.Claims.(jwt.MapClaims)
	if !ok {
		return time.Time{}, false
	}
	rawExp, ok := claims["exp"]
	if !ok {
		return time.Time{}, false
	}

	switch exp := rawExp.(type) {
	case float64:
		return time.Unix(int64(exp), 0), true
	case int64:
		return time.Unix(exp, 0), true
	case int:
		return time.Unix(int64(exp), 0), true
	default:
		return time.Time{}, false
	}
}

func (s *AppState) showError(err error) {
	s.logger.WithError(err).Error("showing error page")
	val, ok := s.components[TvError]
	if !ok {
		s.logger.Fatalf("unable to show error on screen: %v", err)
		return
	}
	val.(*tview.TextView).SetText(err.Error())

	is401 := isUnauthorizedError(err)
	needsAuth := is401 || isAuthMissingError(err)
	if body, ok := s.components[FlErrorBody].(*tview.Flex); ok {
		if action, ok := s.components[FlErrorAction]; ok {
			if needsAuth {
				body.ResizeItem(action, 3, 0)
			} else {
				body.ResizeItem(action, 0, 0)
			}
		}
	}
	s.pages.SwitchToPage(PageError)
	if needsAuth {
		if btn, ok := s.components[BtErrorAuth].(*tview.Button); ok {
			s.app.SetFocus(btn)
		}
	}
	s.app.Draw()
}

func (s *AppState) createErrorView() tview.Primitive {
	p := tview.NewTextView()
	p.SetTitle("ERROR")
	p.SetText("An error has occurred")
	p.SetBackgroundColor(tcell.ColorRed)
	p.SetTextAlign(tview.AlignCenter)
	p.SetBorder(true)
	p.SetBorderPadding(1, 1, 1, 1)

	authButton := tview.NewButton("Run auth refresh")
	authButton.SetSelectedFunc(func() {
		go func() {
			termNote := termEverythingNote()
			s.app.QueueUpdateDraw(func() {
				if termNote == "" {
					p.SetText("Refreshing auth...")
					return
				}
				p.SetText(fmt.Sprintf("Refreshing auth... (%s)", termNote))
			})
			err := s.refreshAuthFromTeamsToken()
			s.app.QueueUpdateDraw(func() {
				if err != nil {
					p.SetText(fmt.Sprintf("Unable to refresh auth: %v", err))
					return
				}
				p.SetText("Auth refresh succeeded. Returning to main view.")
				s.pages.SwitchToPage(PageMain)
				if tree, ok := s.components[TrChat]; ok {
					s.app.SetFocus(tree)
				}
			})
		}()
	})
	actionRow := tview.NewFlex().
		AddItem(nil, 0, 1, false).
		AddItem(authButton, 24, 0, true).
		AddItem(nil, 0, 1, false)

	body := tview.NewFlex().SetDirection(tview.FlexRow).
		AddItem(nil, 0, 1, false).
		AddItem(p, 10, 1, false).
		AddItem(actionRow, 0, 0, false).
		AddItem(nil, 0, 1, false)

	s.components[TvError] = p
	s.components[BtErrorAuth] = authButton
	s.components[FlErrorBody] = body
	s.components[FlErrorAction] = actionRow

	return tview.NewFlex().
		AddItem(nil, 0, 1, false).
		AddItem(body, 60, 1, false).
		AddItem(nil, 0, 1, false)
}

func (s *AppState) fillMainWindow() {
	s.logger.Debug("building main window tree")
	treeView := s.components[TrChat].(*tview.TreeView)
	composeView := s.components[ViCompose].(*tview.InputField)
	rootNode := tview.NewTreeNode("Conversations")
	teamsNode := tview.NewTreeNode("Teams")
	teamsNode.SetColor(tcell.ColorBlue)
	chatsNode := tview.NewTreeNode("Chats")
	chatsNode.SetColor(tcell.ColorYellow)
	favoritesNode := tview.NewTreeNode("Favorites")
	favoritesNode.SetColor(tcell.ColorYellow)
	recentNode := tview.NewTreeNode("Recent")
	recentNode.SetColor(tcell.ColorYellow)

	var firstNode *tview.TreeNode
	var mostRecentChatNode *tview.TreeNode
	for _, t := range s.conversations.Teams {
		currentTeamTreeNode := tview.NewTreeNode(t.DisplayName)
		currentTeamTreeNode.SetReference(t)
		if firstNode == nil {
			firstNode = currentTeamTreeNode
		}

		for _, c := range t.Channels {
			currentChannelTreeNode := tview.NewTreeNode(c.DisplayName)
			currentChannelTreeNode.SetReference(c)
			currentChannelTreeNode.SetColor(tcell.ColorGreen)
			currentTeamTreeNode.AddChild(currentChannelTreeNode)
		}
		currentTeamTreeNode.CollapseAll()
		currentTeamTreeNode.SetColor(tcell.ColorBlue)

		teamsNode.AddChild(currentTeamTreeNode)
	}
	rootNode.AddChild(teamsNode)
	s.logger.WithField("teams_count", len(s.conversations.Teams)).Debug("teams tree nodes prepared")

	chats := append([]csa.Chat(nil), s.conversations.Chats...)
	chats = ensurePrivateNotesChat(chats, s.conversations.PrivateFeeds)
	sort.Slice(chats, func(i, j int) bool {
		ti := chatLastActivity(chats[i])
		tj := chatLastActivity(chats[j])
		if !ti.Equal(tj) {
			return ti.After(tj)
		}
		return buildChatDisplayName(chats[i], s.me) < buildChatDisplayName(chats[j], s.me)
	})
	for _, chat := range chats {
		if strings.TrimSpace(chat.Id) == "" {
			s.logger.Debug("skipping chat with empty id")
			continue
		}
		chatName := buildChatDisplayName(chat, s.me)
		candidateIDs := candidateConversationIds(chat, s.conversations.PrivateFeeds)
		chatKey := chatFavoriteKey(chat.Id, candidateIDs)
		chatName = s.chatDisplayNameForKey(chatKey, chatName)
		s.logger.WithFields(logrus.Fields{
			"chat_title":      chatName,
			"chat_id":         chat.Id,
			"is_one_on_one":   chat.IsOneOnOne,
			"candidate_ids":   strings.Join(candidateIDs, ","),
			"member_count":    len(chat.Members),
			"last_container":  chat.LastMessage.ContainerId,
			"last_message_id": chat.LastMessage.Id,
		}).Debug("prepared chat node")
		chatNode := tview.NewTreeNode(chatName)
		chatNode.SetColor(tcell.ColorGreen)
		isFavorite := s.chatIsFavorite(chat, chatKey)
		isUnread := !chat.IsRead
		if override, ok := s.getManualUnreadOverride(chatKey); ok {
			isUnread = override
		}
		chatNode.SetText(formatChatTreeTitle(chatName, isUnread))
		chatNode.SetReference(conversationRef{
			ids:        candidateIDs,
			title:      chatName,
			chatKey:    chatKey,
			isFavorite: isFavorite,
			isUnread:   isUnread,
		})
		if firstNode == nil {
			firstNode = chatNode
		}
		if mostRecentChatNode == nil {
			mostRecentChatNode = chatNode
		}
		if isFavorite {
			favoritesNode.AddChild(chatNode)
			continue
		}
		recentNode.AddChild(chatNode)
	}
	if len(favoritesNode.GetChildren()) > 0 {
		chatsNode.AddChild(favoritesNode)
	}
	if len(recentNode.GetChildren()) > 0 {
		chatsNode.AddChild(recentNode)
	}
	rootNode.AddChild(chatsNode)
	settingsNode := tview.NewTreeNode("Settings & Help")
	settingsNode.SetColor(tcell.ColorLightSkyBlue)
	settingsNode.SetReference(conversationRef{
		ids:        []string{},
		title:      "Settings & Help",
		chatKey:    settingsHelpChatKey,
		isFavorite: false,
		isUnread:   false,
	})
	rootNode.AddChild(settingsNode)
	s.logger.WithFields(logrus.Fields{
		"chat_nodes_count":      len(favoritesNode.GetChildren()) + len(recentNode.GetChildren()),
		"favorites_nodes_count": len(favoritesNode.GetChildren()),
		"recent_nodes_count":    len(recentNode.GetChildren()),
	}).Debug("chat tree nodes prepared")

	treeView.SetSelectedFunc(func(node *tview.TreeNode) {
		defer func() {
			if recovered := recover(); recovered != nil {
				s.logger.WithFields(logrus.Fields{
					"panic": recovered,
					"stack": string(debug.Stack()),
				}).Error("panic in tree selected handler")
			}
		}()

		s.logger.WithField("node_text", node.GetText()).Debug("tree node selected")
		reference := node.GetReference()
		if reference == nil {
			node.SetExpanded(!node.IsExpanded())
			return
		}

		children := node.GetChildren()
		if len(children) > 0 {
			// Collapse if visible, expand if collapsed.
			node.SetExpanded(!node.IsExpanded())
			return
		}

		switch ref := reference.(type) {
		case csa.Channel:
			channelRef := ref
			s.logger.WithFields(logrus.Fields{
				"target":          "team-channel",
				"display_name":    channelRef.DisplayName,
				"conversation_id": channelRef.Id,
			}).Info("loading conversation")
			s.components[ViChat].(*tview.List).
				SetTitle(channelRef.DisplayName).
				SetBorder(true).
				SetTitleAlign(tview.AlignCenter)
			s.setActiveConversation(node, []string{channelRef.Id}, channelRef.DisplayName)
			go s.loadConversations(&channelRef)
		case conversationRef:
			if ref.chatKey == settingsHelpChatKey {
				s.showSettingsHelpChat(ref.title)
				return
			}
			s.logger.WithFields(logrus.Fields{
				"target":        "chat",
				"display_name":  ref.title,
				"candidate_ids": strings.Join(ref.ids, ","),
			}).Info("loading conversation")
			if ref.isUnread {
				ref.isUnread = false
				s.setManualUnread(ref.chatKey, false)
				node.SetText(formatChatTreeTitle(ref.title, false))
				node.SetReference(ref)
			}
			s.components[ViChat].(*tview.List).
				SetTitle(ref.title).
				SetBorder(true).
				SetTitleAlign(tview.AlignCenter)
			s.setActiveConversation(node, ref.ids, ref.title)
			go s.loadConversationsByIDs(node, ref.ids, ref.title)
		}
	})
	treeView.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		defer func() {
			if recovered := recover(); recovered != nil {
				s.logger.WithFields(logrus.Fields{
					"panic": recovered,
					"stack": string(debug.Stack()),
				}).Error("panic in tree input handler")
			}
		}()

		if s.bindingMatches(actionMoveDown, event) {
			return tcell.NewEventKey(tcell.KeyDown, 0, event.Modifiers())
		}
		if s.bindingMatches(actionMoveUp, event) {
			return tcell.NewEventKey(tcell.KeyUp, 0, event.Modifiers())
		}
		if s.bindingMatches(actionToggleScan, event) {
			enabled := s.toggleUnreadScanEnabled()
			s.logger.WithField("enabled", enabled).Info("unread scan toggle changed")
			s.updateScanStatusTitle()
			composeView.SetTitle(s.composeTitleWithScanStatus())
			return nil
		}
		if s.bindingMatches(actionScanNow, event) {
			if s.markUnreadScanStart() {
				composeView.SetTitle(s.composeTitleWithScanStatus())
				go s.refreshUnreadMarkers(chatsNode)
			} else {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Scan already running")
			}
			return nil
		}
		if s.bindingMatches(actionMarkUnread, event) {
			selected := treeView.GetCurrentNode()
			if selected == nil {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Select a chat first")
				return event
			}
			ref, ok := selected.GetReference().(conversationRef)
			if !ok || strings.TrimSpace(ref.chatKey) == "" || ref.chatKey == settingsHelpChatKey {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Select a chat first")
				return event
			}
			ref.isUnread = true
			s.setManualUnread(ref.chatKey, true)
			selected.SetText(formatChatTreeTitle(ref.title, true))
			selected.SetReference(ref)
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | Marked unread")
			s.logger.WithFields(logrus.Fields{
				"chat_key": ref.chatKey,
				"title":    ref.title,
			}).Debug("marked chat unread manually")
			return nil
		}
		if s.bindingMatches(actionToggleFavorite, event) {
			if s.toggleFavoriteForCurrentNode(treeView, chatsNode, favoritesNode, recentNode) {
				return nil
			}
			return event
		}
		if s.bindingMatches(actionRefreshTitles, event) {
			go s.refreshAllChatLabels(chatsNode)
			return nil
		}
		if s.bindingMatches(actionReloadKeybinds, event) {
			if err := s.reloadKeybindingsConfig(); err != nil {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Keybind reload failed")
			} else {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Keybindings reloaded")
			}
			return nil
		}
		if s.bindingMatches(actionFocusCompose, event) {
			ids, _, _ := s.getActiveConversation()
			if len(ids) == 0 {
				return event
			}
			s.app.SetFocus(composeView)
			return nil
		}
		return event
	})
	chatView := s.components[ViChat].(*tview.List)
	chatView.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		defer func() {
			if recovered := recover(); recovered != nil {
				s.logger.WithFields(logrus.Fields{
					"panic": recovered,
					"stack": string(debug.Stack()),
				}).Error("panic in chat input handler")
			}
		}()

		if s.isSettingsMode() {
			if s.bindingMatches(actionMoveDown, event) {
				return tcell.NewEventKey(tcell.KeyDown, 0, event.Modifiers())
			}
			if s.bindingMatches(actionMoveUp, event) {
				return tcell.NewEventKey(tcell.KeyUp, 0, event.Modifiers())
			}
			captureAction := s.getSettingsCaptureAction()
			if captureAction != "" {
				composeView := s.components[ViCompose].(*tview.InputField)
				if event.Key() == tcell.KeyEscape {
					if err := s.resetActionBindingToDefault(captureAction); err != nil {
						composeView.SetTitle(s.composeTitleWithScanStatus() + " | Reset failed")
					} else {
						composeView.SetTitle(s.composeTitleWithScanStatus() + " | Reset to preset default")
					}
					s.clearSettingsCaptureAction()
					s.renderSettingsHelpItems(chatView)
					return nil
				}
				token := eventToKeybindToken(event)
				if token == "" {
					composeView.SetTitle(s.composeTitleWithScanStatus() + " | Unsupported key (use config for complex)")
					return nil
				}
				if err := s.setActionBinding(captureAction, token); err != nil {
					composeView.SetTitle(s.composeTitleWithScanStatus() + " | Save binding failed")
				} else {
					composeView.SetTitle(s.composeTitleWithScanStatus() + " | Bound " + captureAction + " -> " + token)
				}
				s.clearSettingsCaptureAction()
				s.renderSettingsHelpItems(chatView)
				return nil
			}
			if s.bindingMatches(actionReloadKeybinds, event) {
				if err := s.reloadKeybindingsConfig(); err != nil {
					composeView.SetTitle(s.composeTitleWithScanStatus() + " | Keybind reload failed")
				} else {
					composeView.SetTitle(s.composeTitleWithScanStatus() + " | Keybindings reloaded")
				}
				s.renderSettingsHelpItems(chatView)
				return nil
			}
			return event
		}

		if s.bindingMatches(actionMoveDown, event) {
			return tcell.NewEventKey(tcell.KeyDown, 0, event.Modifiers())
		}
		if s.bindingMatches(actionMoveUp, event) {
			return tcell.NewEventKey(tcell.KeyUp, 0, event.Modifiers())
		}

		if s.bindingMatches(actionReplyMessage, event) {
			current := chatView.GetCurrentItem()
			if current < 0 {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Select a message first")
				return nil
			}
			msg, ok := s.getCurrentChatMessage(current)
			if !ok {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Select a message first")
				return nil
			}
			author := strings.TrimSpace(msg.ImDisplayName)
			if author == "" {
				author = inferMessageAuthor(msg, s.me)
			}
			reply := &replyTarget{
				MessageID: strings.TrimSpace(msg.Id),
				Author:    author,
				Preview:   summarizeReplyPreview(msg.Content),
			}
			s.setPendingReply(reply)
			s.updateComposeReplyUI()
			s.app.SetFocus(composeView)
			return nil
		}
		if s.bindingMatches(actionReactMessage, event) {
			current := chatView.GetCurrentItem()
			if current < 0 {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Select a message first")
				return nil
			}
			msg, ok := s.getCurrentChatMessage(current)
			if !ok || strings.TrimSpace(msg.Id) == "" || strings.TrimSpace(msg.ConversationId) == "" {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Select a message first")
				return nil
			}
			go s.reactToMessage(msg, defaultReactionKey)
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | Reacted 👍")
			return nil
		}
		if s.bindingMatches(actionReloadKeybinds, event) {
			if err := s.reloadKeybindingsConfig(); err != nil {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Keybind reload failed")
			} else {
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | Keybindings reloaded")
			}
			return nil
		}
		return event
	})
	composeView.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		if event == nil {
			return event
		}
		if event.Key() != tcell.KeyUp && event.Key() != tcell.KeyDown {
			return event
		}
		text := composeView.GetText()
		trimmed := strings.TrimRight(text, " \t")
		suffix := text[len(trimmed):]
		start, prefix, query, ok := extractTrailingMentionQuery(trimmed)
		if !ok {
			s.resetMentionCycle()
			return event
		}
		ids, _, _ := s.getActiveConversation()
		tokenKey := strings.ToLower(prefix + ":" + query)
		suggestions, cycleKey := s.getMentionCycleSuggestions(prefix, query, tokenKey, trimmed, ids)
		if len(suggestions) == 0 {
			s.resetMentionCycle()
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | No mention matches")
			return nil
		}

		nextIdx := s.nextMentionCycleIndex(cycleKey, len(suggestions), event.Key() == tcell.KeyUp)
		if nextIdx < 0 || nextIdx >= len(suggestions) {
			return event
		}
		selected := suggestions[nextIdx]
		if strings.TrimSpace(selected.DisplayName) == "" {
			return event
		}
		replacement := prefix + mentionTokenFromDisplayName(selected.DisplayName)
		if strings.TrimSpace(replacement) == strings.TrimSpace(prefix) {
			return event
		}
		composeView.SetText(trimmed[:start] + replacement + suffix)
		s.setMentionCycleItems(cycleKey, suggestions)
		composeView.SetTitle(s.composeTitleWithScanStatus() + " | Mention: " + selected.DisplayName)
		return nil
	})
	composeView.SetDoneFunc(func(key tcell.Key) {
		if key == tcell.KeyEscape {
			s.resetMentionCycle()
			s.clearPendingReply()
			s.updateComposeReplyUI()
			s.app.SetFocus(treeView)
			return
		}
		if key != tcell.KeyEnter {
			return
		}
		messageText := strings.TrimSpace(composeView.GetText())
		if messageText == "" {
			return
		}
		ids, title, selectedNode := s.getActiveConversation()
		if len(ids) == 0 {
			s.showError(fmt.Errorf("select a conversation before sending a message"))
			return
		}
		reply := s.getPendingReply()
		s.clearPendingReply()
		s.resetMentionCycle()
		s.updateComposeReplyUI()
		composeView.SetText("")
		go s.sendMessageAndRefresh(ids, title, messageText, selectedNode, reply)
	})

	treeView.SetRoot(rootNode)
	if mostRecentChatNode != nil {
		treeView.SetCurrentNode(mostRecentChatNode)
		if ref, ok := mostRecentChatNode.GetReference().(conversationRef); ok {
			s.components[ViChat].(*tview.List).
				SetTitle(ref.title).
				SetBorder(true).
				SetTitleAlign(tview.AlignCenter)
			s.setActiveConversation(mostRecentChatNode, ref.ids, ref.title)
			go s.loadConversationsByIDs(mostRecentChatNode, ref.ids, ref.title)
		}
	} else if firstNode != nil {
		treeView.SetCurrentNode(firstNode)
	} else {
		treeView.SetCurrentNode(rootNode)
	}

	s.pages.SwitchToPage(PageMain)
	s.app.SetFocus(treeView)
	s.app.Draw()
	s.startUnreadScanLoop(chatsNode)
	s.logger.Info("main window ready")
}

func (s *AppState) setActiveConversation(selectedNode *tview.TreeNode, conversationIDs []string, title string) {
	ids := normalizeConversationIDs(conversationIDs)
	s.activeConversationMu.Lock()
	s.activeConversationIDs = ids
	s.activeConversationTitle = title
	s.activeConversationNode = selectedNode
	s.activeConversationMu.Unlock()
	s.setSettingsMode(false)
	s.clearPendingReply()
	s.updateComposeReplyUI()
}

func (s *AppState) showSettingsHelpChat(title string) {
	chatList := s.components[ViChat].(*tview.List)
	chatList.Clear()
	chatList.SetTitle(title).
		SetBorder(true).
		SetTitleAlign(tview.AlignCenter)
	s.renderSettingsHelpItems(chatList)
	chatList.SetChangedFunc(func(index int, _ string, _ string, _ rune) {
		s.setSettingsSelection(index)
	})
	chatList.SetSelectedFunc(func(index int, _ string, _ string, _ rune) {
		s.handleSettingsSelection(index)
	})

	s.setCurrentChatMessages(nil)
	s.setActiveConversation(nil, nil, title)
	s.settingsMu.Lock()
	s.settingsMode = true
	s.settingsSelection = 0
	s.settingsCaptureAction = ""
	s.settingsMu.Unlock()
	s.logger.Debug("settings/help chat rendered")
}

func (s *AppState) renderSettingsHelpItems(chatList *tview.List) {
	if chatList == nil {
		return
	}
	selection := s.getSettingsSelection()
	items := s.buildSettingsItems()
	chatList.Clear()
	for _, item := range items {
		switch item.kind {
		case settingsItemSpacer:
			chatList.AddItem(" ", "", 0, nil)
		case settingsItemOpen:
			chatList.AddItem("Open Keybindings Config", s.keybindPath+" (Enter)", 0, nil)
		case settingsItemPreset:
			chatList.AddItem("Preset", s.formatPresetLine()+" (Enter to cycle)", 0, nil)
		case settingsItemWrap:
			chatList.AddItem("Chat Text Mode", s.formatChatWrapLine()+" (Enter to toggle)", 0, nil)
		case settingsItemWrapPct:
			chatList.AddItem("Wrap Chars", s.formatChatWrapPctLine()+" (Enter to cycle)", 0, nil)
		case settingsItemComposeColor:
			chatList.AddItem("Compose Color", s.formatComposeColorLine()+" (Enter to cycle)", 0, nil)
		case settingsItemAuthorColor:
			chatList.AddItem("Username Color", s.formatAuthorColorLine()+" (Enter to cycle)", 0, nil)
		case settingsItemReload:
			chatList.AddItem("Reload Keybindings", "Reload from config file (Enter/Ctrl+R)", 0, nil)
		case settingsItemBinding:
			chatList.AddItem("Bind "+item.action, s.formatActionBindingLine(item.action)+" (Enter to rebind)", 0, nil)
		default:
			chatList.AddItem("Help", "Esc in bind mode resets that action to preset default", 0, nil)
		}
	}
	if selection >= 0 && selection < chatList.GetItemCount() {
		chatList.SetCurrentItem(selection)
	}
}

func (s *AppState) buildSettingsItems() []settingsItem {
	return []settingsItem{
		{kind: settingsItemOpen},
		{kind: settingsItemPreset},
		{kind: settingsItemSpacer},
		{kind: settingsItemWrap},
		{kind: settingsItemWrapPct},
		{kind: settingsItemComposeColor},
		{kind: settingsItemAuthorColor},
		{kind: settingsItemSpacer},
		{kind: settingsItemReload},
		{kind: settingsItemSpacer},
		{kind: settingsItemBinding, action: actionMoveDown},
		{kind: settingsItemBinding, action: actionMoveUp},
		{kind: settingsItemBinding, action: actionFocusCompose},
		{kind: settingsItemBinding, action: actionToggleFavorite},
		{kind: settingsItemBinding, action: actionMarkUnread},
		{kind: settingsItemBinding, action: actionReplyMessage},
		{kind: settingsItemBinding, action: actionReactMessage},
		{kind: settingsItemBinding, action: actionRefreshTitles},
		{kind: settingsItemBinding, action: actionToggleScan},
		{kind: settingsItemBinding, action: actionScanNow},
		{kind: settingsItemBinding, action: actionReloadKeybinds},
		{kind: settingsItemSpacer},
		{kind: settingsItemInfo},
	}
}

func (s *AppState) handleSettingsSelection(index int) {
	items := s.buildSettingsItems()
	if index < 0 || index >= len(items) {
		return
	}
	s.setSettingsSelection(index)
	item := items[index]
	composeView := s.components[ViCompose].(*tview.InputField)
	switch item.kind {
	case settingsItemSpacer, settingsItemInfo:
		return
	case settingsItemOpen:
		err := s.openKeybindConfigInEditor()
		if err != nil {
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | Unable to open editor")
		} else {
			_ = s.reloadKeybindingsConfig()
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | Editor closed, keybindings reloaded")
		}
		s.renderSettingsHelpItems(s.components[ViChat].(*tview.List))
	case settingsItemPreset:
		next := nextPreset(s.getCurrentKeybindPreset())
		if err := s.setKeybindPreset(next); err != nil {
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | Preset change failed")
		} else {
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | Preset: " + next)
		}
		s.renderSettingsHelpItems(s.components[ViChat].(*tview.List))
	case settingsItemReload:
		if err := s.reloadKeybindingsConfig(); err != nil {
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | Keybind reload failed")
		} else {
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | Keybindings reloaded")
		}
		s.renderSettingsHelpItems(s.components[ViChat].(*tview.List))
	case settingsItemWrap:
		s.toggleChatWordWrap()
		s.persistEncryptedChatSettings()
		s.rerenderActiveChatMessages()
		composeView.SetTitle(s.composeTitleWithScanStatus() + " | " + s.formatChatWrapLine())
		s.renderSettingsHelpItems(s.components[ViChat].(*tview.List))
	case settingsItemWrapPct:
		next, promptCustom := s.cycleChatWrapPercent()
		if promptCustom {
			s.promptCustomWrapChars(func(value int, ok bool) {
				if !ok {
					composeView.SetTitle(s.composeTitleWithScanStatus() + " | Custom wrap canceled")
					s.renderSettingsHelpItems(s.components[ViChat].(*tview.List))
					return
				}
				s.setChatWrapChars(value)
				s.persistEncryptedChatSettings()
				s.rerenderActiveChatMessages()
				composeView.SetTitle(s.composeTitleWithScanStatus() + " | " + s.formatChatWrapPctLine())
				s.renderSettingsHelpItems(s.components[ViChat].(*tview.List))
			})
		} else {
			s.setChatWrapChars(next)
			s.persistEncryptedChatSettings()
			s.rerenderActiveChatMessages()
			composeView.SetTitle(s.composeTitleWithScanStatus() + " | " + s.formatChatWrapPctLine())
			s.renderSettingsHelpItems(s.components[ViChat].(*tview.List))
		}
	case settingsItemComposeColor:
		s.cycleComposeColor()
		s.applyComposeTheme()
		s.persistEncryptedChatSettings()
		composeView.SetTitle(s.composeTitleWithScanStatus() + " | Compose: " + s.formatComposeColorLine())
		s.renderSettingsHelpItems(s.components[ViChat].(*tview.List))
	case settingsItemAuthorColor:
		s.cycleAuthorColor()
		s.persistEncryptedChatSettings()
		s.rerenderActiveChatMessages()
		composeView.SetTitle(s.composeTitleWithScanStatus() + " | Usernames: " + s.formatAuthorColorLine())
		s.renderSettingsHelpItems(s.components[ViChat].(*tview.List))
	case settingsItemBinding:
		s.settingsMu.Lock()
		s.settingsCaptureAction = item.action
		s.settingsMu.Unlock()
		composeView.SetTitle(s.composeTitleWithScanStatus() + " | Press new key or Esc for default")
	}
}

func nextPreset(current string) string {
	order := []string{"default", "vim", "emacs", "jk"}
	current = strings.ToLower(strings.TrimSpace(current))
	for i, v := range order {
		if v == current {
			return order[(i+1)%len(order)]
		}
	}
	return order[0]
}

func (s *AppState) getCurrentKeybindPreset() string {
	s.keybindMu.RLock()
	defer s.keybindMu.RUnlock()
	return s.keybindPreset
}

func (s *AppState) formatPresetLine() string {
	current := strings.ToLower(strings.TrimSpace(s.getCurrentKeybindPreset()))
	all := []string{"default", "vim", "emacs", "jk"}
	parts := make([]string, 0, len(all))
	for _, p := range all {
		if p == current {
			parts = append(parts, "["+p+"]")
		} else {
			parts = append(parts, p)
		}
	}
	return strings.Join(parts, " ")
}

func (s *AppState) formatActionBindingLine(action string) string {
	s.keybindMu.RLock()
	keys := append([]string(nil), s.keybindings[action]...)
	s.keybindMu.RUnlock()
	if len(keys) == 0 {
		return "(none)"
	}
	return strings.Join(keys, ", ")
}

func (s *AppState) composeFieldColor() tcell.Color {
	s.themeMu.RLock()
	name := s.composeColorName
	s.themeMu.RUnlock()
	switch strings.ToLower(strings.TrimSpace(name)) {
	case "midnight":
		return tcell.ColorMidnightBlue
	case "navy":
		return tcell.ColorNavy
	case "dark_blue":
		return tcell.ColorDarkBlue
	case "slate":
		return tcell.ColorDarkSlateBlue
	default:
		return tcell.ColorDarkBlue
	}
}

func (s *AppState) authorStyleTag() string {
	s.themeMu.RLock()
	name := strings.ToLower(strings.TrimSpace(s.authorColorName))
	s.themeMu.RUnlock()
	switch name {
	case "yellow":
		return "yellow"
	case "green":
		return "green"
	case "cyan":
		return "cyan"
	case "white":
		return "white"
	case "blue":
		return "blue"
	default:
		return "blue"
	}
}

func (s *AppState) cycleComposeColor() string {
	choices := []string{"dark_blue", "midnight", "navy", "slate"}
	s.themeMu.Lock()
	defer s.themeMu.Unlock()
	current := strings.ToLower(strings.TrimSpace(s.composeColorName))
	idx := 0
	for i, c := range choices {
		if c == current {
			idx = i
			break
		}
	}
	next := choices[(idx+1)%len(choices)]
	s.composeColorName = next
	return next
}

func (s *AppState) cycleAuthorColor() string {
	choices := []string{"blue", "yellow", "green", "cyan", "white"}
	s.themeMu.Lock()
	defer s.themeMu.Unlock()
	current := strings.ToLower(strings.TrimSpace(s.authorColorName))
	idx := 0
	for i, c := range choices {
		if c == current {
			idx = i
			break
		}
	}
	next := choices[(idx+1)%len(choices)]
	s.authorColorName = next
	return next
}

func (s *AppState) formatComposeColorLine() string {
	s.themeMu.RLock()
	name := s.composeColorName
	s.themeMu.RUnlock()
	return strings.ToLower(strings.TrimSpace(name))
}

func (s *AppState) formatAuthorColorLine() string {
	s.themeMu.RLock()
	name := s.authorColorName
	s.themeMu.RUnlock()
	return strings.ToLower(strings.TrimSpace(name))
}

func (s *AppState) applyComposeTheme() {
	val, ok := s.components[ViCompose]
	if !ok {
		return
	}
	input, ok := val.(*tview.InputField)
	if !ok {
		return
	}
	input.SetFieldBackgroundColor(s.composeFieldColor())
	input.SetFieldTextColor(tcell.ColorWhite)
	input.SetLabelColor(tcell.ColorWhite)
	input.SetPlaceholderTextColor(tcell.ColorLightGray)
}

func (s *AppState) isChatWordWrap() bool {
	s.chatWordWrapMu.RLock()
	defer s.chatWordWrapMu.RUnlock()
	return s.chatWordWrap
}

func (s *AppState) toggleChatWordWrap() bool {
	s.chatWordWrapMu.Lock()
	s.chatWordWrap = !s.chatWordWrap
	v := s.chatWordWrap
	s.chatWordWrapMu.Unlock()
	return v
}

func (s *AppState) formatChatWrapLine() string {
	if s.isChatWordWrap() {
		return "[Word Wrap] Scroll"
	}
	return "Word Wrap [Scroll]"
}

func (s *AppState) getChatWrapPercent() int {
	s.chatWordWrapMu.RLock()
	defer s.chatWordWrapMu.RUnlock()
	if s.chatWrapChars <= 0 {
		return 80
	}
	return s.chatWrapChars
}

func (s *AppState) formatChatWrapPctLine() string {
	s.chatWordWrapMu.RLock()
	chars := s.chatWrapChars
	effective := s.chatWrapEffective
	s.chatWordWrapMu.RUnlock()
	if chars <= 0 {
		chars = 80
	}
	if effective > 0 && effective != chars {
		return fmt.Sprintf("%d chars (effective %d)", chars, effective)
	}
	return fmt.Sprintf("%d chars", chars)
}

func (s *AppState) cycleChatWrapPercent() (int, bool) {
	choices := []int{20, 40, 72, 80, 100, 200, 400, 600, 800, 1000}
	s.chatWordWrapMu.Lock()
	defer s.chatWordWrapMu.Unlock()
	current := s.chatWrapChars
	if current <= 0 {
		current = 80
	}
	idx := -1
	for i, v := range choices {
		if v == current {
			idx = i
			break
		}
	}
	if idx == len(choices)-1 {
		return current, true
	}
	if idx < 0 {
		return choices[0], false
	}
	return choices[(idx+1)%len(choices)], false
}

func (s *AppState) setChatWrapChars(v int) {
	if v < 8 {
		v = 8
	}
	if v > 5000 {
		v = 5000
	}
	s.chatWordWrapMu.Lock()
	s.chatWrapChars = v
	s.chatWordWrapMu.Unlock()
}

func (s *AppState) promptCustomWrapChars(onDone func(value int, ok bool)) {
	defaultValue := strconv.Itoa(s.getChatWrapPercent())
	input := tview.NewInputField().
		SetLabel("Wrap chars: ").
		SetText(defaultValue)
	input.SetFieldWidth(12)
	input.SetAcceptanceFunc(func(text string, lastChar rune) bool {
		if lastChar == 0 {
			return true
		}
		return lastChar >= '0' && lastChar <= '9'
	})

	form := tview.NewForm().
		AddFormItem(input).
		AddButton("Save", func() {
			value, err := strconv.Atoi(strings.TrimSpace(input.GetText()))
			if err != nil || value <= 0 {
				if onDone != nil {
					onDone(0, false)
				}
			} else {
				if onDone != nil {
					onDone(value, true)
				}
			}
			s.pages.RemovePage("pageWrapCustom")
			s.pages.SwitchToPage(PageMain)
			if chat, ok := s.components[ViChat]; ok {
				s.app.SetFocus(chat)
			}
		}).
		AddButton("Cancel", func() {
			if onDone != nil {
				onDone(0, false)
			}
			s.pages.RemovePage("pageWrapCustom")
			s.pages.SwitchToPage(PageMain)
			if chat, ok := s.components[ViChat]; ok {
				s.app.SetFocus(chat)
			}
		})
	form.SetBorder(true).SetTitle("Custom Wrap Width").SetTitleAlign(tview.AlignCenter)

	modal := tview.NewFlex().
		AddItem(nil, 0, 1, false).
		AddItem(tview.NewFlex().SetDirection(tview.FlexRow).
			AddItem(nil, 0, 1, false).
			AddItem(form, 7, 1, true).
			AddItem(nil, 0, 1, false), 60, 1, true).
		AddItem(nil, 0, 1, false)

	s.pages.AddPage("pageWrapCustom", modal, true, true)
	s.app.SetFocus(input)
}

func (s *AppState) formatChatMessageText(content string) string {
	txt := textMessage(content)
	if s.isChatWordWrap() {
		return txt
	}
	return strings.Join(strings.Fields(txt), " ")
}

func wrapTextLines(text string, width int) []string {
	if strings.TrimSpace(text) == "" {
		return []string{""}
	}
	if width <= 8 {
		width = 80
	}
	paras := strings.Split(text, "\n")
	lines := []string{}
	for _, p := range paras {
		p = strings.TrimRight(p, " \t\r")
		if strings.TrimSpace(p) == "" {
			lines = append(lines, "")
			continue
		}
		rest := p
		for len(rest) > 0 {
			if len(rest) <= width {
				lines = append(lines, rest)
				break
			}
			cut := width
			segment := rest[:cut]
			if idx := strings.LastIndex(segment, " "); idx > 0 {
				cut = idx
			}
			lines = append(lines, strings.TrimRight(rest[:cut], " "))
			rest = strings.TrimLeft(rest[cut:], " ")
		}
	}
	if len(lines) == 0 {
		return []string{""}
	}
	return lines
}

func (s *AppState) rerenderActiveChatMessages() {
	ids, title, node := s.getActiveConversation()
	if len(ids) == 0 {
		return
	}
	go s.loadConversationsByIDs(node, ids, title)
}

func (s *AppState) isSettingsMode() bool {
	s.settingsMu.RLock()
	defer s.settingsMu.RUnlock()
	return s.settingsMode
}

func (s *AppState) setSettingsMode(v bool) {
	s.settingsMu.Lock()
	s.settingsMode = v
	if !v {
		s.settingsCaptureAction = ""
	}
	s.settingsMu.Unlock()
}

func (s *AppState) getSettingsSelection() int {
	s.settingsMu.RLock()
	defer s.settingsMu.RUnlock()
	return s.settingsSelection
}

func (s *AppState) setSettingsSelection(index int) {
	s.settingsMu.Lock()
	s.settingsSelection = index
	s.settingsMu.Unlock()
}

func (s *AppState) getSettingsCaptureAction() string {
	s.settingsMu.RLock()
	defer s.settingsMu.RUnlock()
	return s.settingsCaptureAction
}

func (s *AppState) clearSettingsCaptureAction() {
	s.settingsMu.Lock()
	s.settingsCaptureAction = ""
	s.settingsMu.Unlock()
}

func (s *AppState) openKeybindConfigInEditor() error {
	editor := strings.TrimSpace(os.Getenv("VISUAL"))
	if editor == "" {
		editor = strings.TrimSpace(os.Getenv("EDITOR"))
	}
	if editor == "" {
		editor = "nano"
	}
	parts := strings.Fields(editor)
	if len(parts) == 0 {
		parts = []string{"nano"}
	}
	args := append(parts[1:], s.keybindPath)
	cmd := exec.Command(parts[0], args...)
	cmd.Stdin = os.Stdin
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	var runErr error
	s.app.Suspend(func() {
		runErr = cmd.Run()
	})
	return runErr
}

func eventToKeybindToken(event *tcell.EventKey) string {
	if event == nil {
		return ""
	}
	switch event.Key() {
	case tcell.KeyUp:
		return "up"
	case tcell.KeyDown:
		return "down"
	case tcell.KeyLeft:
		return "left"
	case tcell.KeyRight:
		return "right"
	case tcell.KeyEnter:
		return "enter"
	case tcell.KeyEscape:
		return "esc"
	case tcell.KeyCtrlN:
		return "ctrl+n"
	case tcell.KeyCtrlP:
		return "ctrl+p"
	case tcell.KeyCtrlX:
		return "ctrl+x"
	case tcell.KeyCtrlR:
		return "ctrl+r"
	case tcell.KeyRune:
		r := event.Rune()
		if r == 0 {
			return ""
		}
		return string(r)
	default:
		return ""
	}
}

func (s *AppState) getActiveConversation() ([]string, string, *tview.TreeNode) {
	s.activeConversationMu.RLock()
	ids := append([]string(nil), s.activeConversationIDs...)
	title := s.activeConversationTitle
	node := s.activeConversationNode
	s.activeConversationMu.RUnlock()
	return ids, title, node
}

func buildChatDisplayName(chat csa.Chat, me *models.User) string {
	if isPrivateNotesChat(chat) {
		return "Private Notes"
	}
	if strings.TrimSpace(chat.Title) != "" {
		return chat.Title
	}

	memberNames := []string{}
	seen := map[string]struct{}{}
	for _, member := range chat.Members {
		if isCurrentUser(member, me) {
			continue
		}
		name := strings.TrimSpace(member.FriendlyName)
		if isSelfDisplayName(name, me) {
			continue
		}
		if name == "" {
			continue
		}
		if _, ok := seen[name]; ok {
			continue
		}
		seen[name] = struct{}{}
		memberNames = append(memberNames, name)
	}
	if len(memberNames) > 0 {
		return strings.Join(memberNames, ", ")
	}
	if chat.IsOneOnOne {
		return "Chat"
	}
	return "Private Chat"
}

func formatChatTreeTitle(title string, unread bool) string {
	if !unread {
		return title
	}
	return "● " + title
}

func (s *AppState) chatIsFavorite(chat csa.Chat, chatKey string) bool {
	key := normalizeFavoriteKey(chatKey)
	if key != "" {
		s.chatFavoritesMu.RLock()
		favorite, ok := s.chatFavorites[key]
		s.chatFavoritesMu.RUnlock()
		if ok {
			return favorite
		}
	}
	return isPrivateNotesChat(chat)
}

func (s *AppState) chatDisplayNameForKey(chatKey, fallback string) string {
	key := normalizeFavoriteKey(chatKey)
	if key == "" {
		return fallback
	}
	s.chatTitlesMu.RLock()
	title := strings.TrimSpace(s.chatTitles[key])
	s.chatTitlesMu.RUnlock()
	if title == "" {
		return fallback
	}
	return title
}

func (s *AppState) toggleChatFavorite(chatKey string, current bool) bool {
	key := normalizeFavoriteKey(chatKey)
	if key == "" {
		return current
	}
	next := !current
	s.chatFavoritesMu.Lock()
	if s.chatFavorites == nil {
		s.chatFavorites = map[string]bool{}
	}
	s.chatFavorites[key] = next
	s.chatFavoritesMu.Unlock()
	s.persistEncryptedChatSettings()
	return next
}

func (s *AppState) setChatTitle(chatKey, title string) bool {
	key := normalizeFavoriteKey(chatKey)
	if key == "" {
		return false
	}
	trimmed := strings.TrimSpace(title)
	if trimmed == "" {
		return false
	}

	s.chatTitlesMu.Lock()
	if s.chatTitles == nil {
		s.chatTitles = map[string]string{}
	}
	if s.chatTitles[key] == trimmed {
		s.chatTitlesMu.Unlock()
		return false
	}
	s.chatTitles[key] = trimmed
	s.chatTitlesMu.Unlock()
	return true
}

func (s *AppState) toggleFavoriteForCurrentNode(treeView *tview.TreeView, chatsNode, favoritesNode, recentNode *tview.TreeNode) (handled bool) {
	defer func() {
		if recovered := recover(); recovered != nil {
			s.logger.WithFields(logrus.Fields{
				"panic": recovered,
				"stack": string(debug.Stack()),
			}).Error("panic while toggling favorite")
			handled = true
		}
	}()

	if treeView == nil {
		return false
	}
	selected := treeView.GetCurrentNode()
	if selected == nil {
		return false
	}
	ref, ok := selected.GetReference().(conversationRef)
	if !ok || strings.TrimSpace(ref.chatKey) == "" {
		return false
	}
	ref.isFavorite = s.toggleChatFavorite(ref.chatKey, ref.isFavorite)
	selected.SetReference(ref)
	moveChatNodeToGroup(chatsNode, favoritesNode, recentNode, selected, ref.isFavorite)
	treeView.SetCurrentNode(selected)
	return true
}

func resolveDMDisplayName(currentName string, messages []csa.ChatMessage, me *models.User) string {
	name := strings.TrimSpace(currentName)
	if !strings.EqualFold(name, "Chat") && !strings.EqualFold(name, "Direct Message") {
		return currentName
	}

	// Prefer the latest non-self author as the DM title.
	for i := len(messages) - 1; i >= 0; i-- {
		author := strings.TrimSpace(messages[i].ImDisplayName)
		if author == "" || isSelfDisplayName(author, me) {
			continue
		}
		return author
	}

	authorCounts := map[string]int{}
	for _, message := range messages {
		author := strings.TrimSpace(message.ImDisplayName)
		if author == "" || isSelfDisplayName(author, me) {
			continue
		}
		authorCounts[author]++
	}

	bestAuthor := ""
	bestCount := 0
	for author, count := range authorCounts {
		if count > bestCount || (count == bestCount && strings.ToLower(author) < strings.ToLower(bestAuthor)) {
			bestAuthor = author
			bestCount = count
		}
	}
	if bestAuthor != "" {
		return bestAuthor
	}

	return currentName
}

func chatLastActivity(chat csa.Chat) time.Time {
	compose := time.Time(chat.LastMessage.ComposeTime)
	if !compose.IsZero() {
		return compose
	}
	arrival := time.Time(chat.LastMessage.OriginalArrivalTime)
	if !arrival.IsZero() {
		return arrival
	}
	return time.Time{}
}

func isSelfDisplayName(name string, me *models.User) bool {
	if me == nil {
		return false
	}
	name = strings.ToLower(strings.TrimSpace(name))
	if name == "" {
		return false
	}

	candidates := []string{
		me.DisplayName,
		me.GivenName,
		strings.TrimSpace(me.GivenName + " " + me.Surname),
		me.Alias,
		me.Email,
		me.UserPrincipalName,
	}
	for _, candidate := range candidates {
		candidate = strings.ToLower(strings.TrimSpace(candidate))
		if candidate != "" && candidate == name {
			return true
		}
	}
	return false
}

func isCurrentUser(member csa.ChatMember, me *models.User) bool {
	if me == nil {
		return false
	}
	if strings.TrimSpace(me.ObjectId) != "" && member.ObjectId == me.ObjectId {
		return true
	}
	if strings.TrimSpace(me.Mri) != "" && member.Mri == me.Mri {
		return true
	}
	return false
}

func candidateConversationIds(chat csa.Chat, feeds []csa.PrivateFeed) []string {
	candidates := []string{}
	seen := map[string]struct{}{}

	for _, id := range []string{chat.Id, chat.LastMessage.ContainerId} {
		id = strings.TrimSpace(id)
		if id == "" {
			continue
		}
		if _, ok := seen[id]; ok {
			continue
		}
		seen[id] = struct{}{}
		candidates = append(candidates, id)
	}

	for _, feed := range feeds {
		id := strings.TrimSpace(extractConversationID(feed.TargetLink))
		if id == "" {
			continue
		}
		if _, ok := seen[id]; ok {
			continue
		}
		if chat.LastMessage.ContainerId != "" && feed.LastMessage.ContainerId != "" && chat.LastMessage.ContainerId != feed.LastMessage.ContainerId {
			continue
		}
		if chat.Id != "" && feed.Id != "" && !strings.Contains(feed.Id, chat.Id) {
			// keep likely-matching feeds by container id and avoid unrelated links
			if chat.LastMessage.ContainerId == "" || feed.LastMessage.ContainerId == "" {
				continue
			}
		}
		seen[id] = struct{}{}
		candidates = append(candidates, id)
	}

	return candidates
}

func (s *AppState) refreshAllChatLabels(chatsNode *tview.TreeNode) {
	if chatsNode == nil {
		return
	}

	nodes := flattenConversationNodes(chatsNode)
	updated := 0
	attempted := 0
	titlesChanged := false
	for _, node := range nodes {
		ref, ok := node.GetReference().(conversationRef)
		if !ok {
			continue
		}

		attempted++
		title, err := s.resolveConversationTitle(ref.title, ref.ids)
		if err != nil {
			s.logger.WithFields(logrus.Fields{
				"candidate_ids": strings.Join(ref.ids, ","),
				"display_name":  ref.title,
				"error":         err.Error(),
			}).Warn("unable to refresh chat title")
			continue
		}
		if strings.TrimSpace(title) == "" || strings.TrimSpace(title) == strings.TrimSpace(ref.title) {
			continue
		}

		ref.title = title
		updated++
		if s.setChatTitle(ref.chatKey, title) {
			titlesChanged = true
		}
		displayTitle := formatChatTreeTitle(ref.title, ref.isUnread)
		s.app.QueueUpdateDraw(func() {
			node.SetText(displayTitle)
			node.SetReference(ref)
		})
	}
	if titlesChanged {
		s.persistEncryptedChatSettings()
	}

	s.logger.WithFields(logrus.Fields{
		"attempted": attempted,
		"updated":   updated,
	}).Info("finished refreshing chat titles")
}

func (s *AppState) startUnreadScanLoop(chatsNode *tview.TreeNode) {
	if chatsNode == nil {
		return
	}
	interval := s.unreadScanInterval
	if interval <= 0 {
		interval = time.Minute
	}
	ticker := time.NewTicker(interval)
	go func() {
		defer ticker.Stop()
		for {
			select {
			case <-ticker.C:
				s.unreadScanMu.RLock()
				enabled := s.unreadScanEnabled
				s.unreadScanMu.RUnlock()
				if !enabled {
					continue
				}
				if s.markUnreadScanStart() {
					s.refreshUnreadMarkers(chatsNode)
				}
			case <-s.unreadScanStop:
				return
			}
		}
	}()
}

func (s *AppState) composeTitleWithScanStatus() string {
	s.unreadScanMu.RLock()
	enabled := s.unreadScanEnabled
	running := s.unreadScanRunning
	last := s.unreadLastScanAt
	changes := s.unreadLastChanges
	s.unreadScanMu.RUnlock()

	status := "OFF"
	if enabled {
		status = "ON"
	}
	replySuffix := ""
	if reply := s.getPendingReply(); reply != nil {
		replySuffix = " | Reply: " + strings.TrimSpace(reply.Author)
	}
	if running {
		return fmt.Sprintf("Compose | Scan: %s | Scanning...%s", status, replySuffix)
	}
	if !last.IsZero() {
		return fmt.Sprintf("Compose | Scan: %s | Last: %s (%d)%s", status, last.Format("15:04:05"), changes, replySuffix)
	}
	return fmt.Sprintf("Compose | Scan: %s%s", status, replySuffix)
}

func (s *AppState) updateScanStatusTitle() {
	val, ok := s.components[ViCompose]
	if !ok {
		return
	}
	input, ok := val.(*tview.InputField)
	if !ok {
		return
	}
	input.SetTitle(s.composeTitleWithScanStatus())
}

func (s *AppState) updateComposeReplyUI() {
	val, ok := s.components[ViCompose]
	if !ok {
		return
	}
	input, ok := val.(*tview.InputField)
	if !ok {
		return
	}
	if reply := s.getPendingReply(); reply != nil {
		input.SetPlaceholder(fmt.Sprintf("Reply to %s: type message and press Enter", strings.TrimSpace(reply.Author)))
	} else {
		input.SetPlaceholder(composeDefaultPlaceholder)
	}
	input.SetTitle(s.composeTitleWithScanStatus())
}

func (s *AppState) isUnreadScanEnabled() bool {
	s.unreadScanMu.RLock()
	enabled := s.unreadScanEnabled
	s.unreadScanMu.RUnlock()
	return enabled
}

func (s *AppState) toggleUnreadScanEnabled() bool {
	s.unreadScanMu.Lock()
	s.unreadScanEnabled = !s.unreadScanEnabled
	enabled := s.unreadScanEnabled
	s.unreadScanMu.Unlock()
	return enabled
}

func (s *AppState) markUnreadScanStart() bool {
	s.unreadScanMu.Lock()
	if s.unreadScanRunning {
		s.unreadScanMu.Unlock()
		return false
	}
	s.unreadScanRunning = true
	s.unreadScanMu.Unlock()
	return true
}

func (s *AppState) markUnreadScanDone(changes int) {
	s.unreadScanMu.Lock()
	s.unreadScanRunning = false
	s.unreadLastScanAt = time.Now()
	s.unreadLastChanges = changes
	s.unreadScanMu.Unlock()
}

func (s *AppState) refreshUnreadMarkers(chatsNode *tview.TreeNode) {
	defer func() {
		if recovered := recover(); recovered != nil {
			s.logger.WithFields(logrus.Fields{
				"panic": recovered,
				"stack": string(debug.Stack()),
			}).Error("panic during unread marker refresh")
			s.markUnreadScanDone(0)
		}
	}()

	conversations, err := s.teamsClient.GetConversations()
	if err != nil && isUnauthorizedError(err) {
		if refreshErr := s.refreshAuthFromTeamsToken(); refreshErr == nil {
			conversations, err = s.teamsClient.GetConversations()
		}
	}
	if err != nil || conversations == nil {
		if err != nil {
			s.logger.WithError(err).Warn("unable to refresh unread state")
		}
		s.markUnreadScanDone(0)
		return
	}

	chats := ensurePrivateNotesChat(conversations.Chats, conversations.PrivateFeeds)
	unreadByKey := map[string]bool{}
	for _, chat := range chats {
		key := chatFavoriteKey(chat.Id, candidateConversationIds(chat, conversations.PrivateFeeds))
		if key == "" {
			continue
		}
		unread := !chat.IsRead
		if override, ok := s.getManualUnreadOverride(key); ok {
			unread = override
			if unread == !chat.IsRead {
				// Override has been reflected by server state.
				s.clearManualUnreadOverride(key)
			}
		}
		unreadByKey[key] = unread
	}

	s.app.QueueUpdateDraw(func() {
		changed := 0
		nodes := flattenConversationNodes(chatsNode)
		for _, node := range nodes {
			ref, ok := node.GetReference().(conversationRef)
			if !ok || strings.TrimSpace(ref.chatKey) == "" {
				continue
			}
			unread, exists := unreadByKey[normalizeFavoriteKey(ref.chatKey)]
			if !exists || unread == ref.isUnread {
				continue
			}
			ref.isUnread = unread
			node.SetText(formatChatTreeTitle(ref.title, unread))
			node.SetReference(ref)
			changed++
		}
		s.markUnreadScanDone(changed)
		s.updateScanStatusTitle()
	})
}

func (s *AppState) setManualUnread(chatKey string, unread bool) {
	key := normalizeFavoriteKey(chatKey)
	if key == "" {
		return
	}
	s.manualUnreadMu.Lock()
	if s.manualUnread == nil {
		s.manualUnread = map[string]bool{}
	}
	s.manualUnread[key] = unread
	s.manualUnreadMu.Unlock()
	s.persistEncryptedChatSettings()
}

func (s *AppState) clearManualUnreadOverride(chatKey string) {
	key := normalizeFavoriteKey(chatKey)
	if key == "" {
		return
	}
	s.manualUnreadMu.Lock()
	if s.manualUnread != nil {
		delete(s.manualUnread, key)
	}
	s.manualUnreadMu.Unlock()
}

func (s *AppState) getManualUnreadOverride(chatKey string) (bool, bool) {
	key := normalizeFavoriteKey(chatKey)
	if key == "" {
		return false, false
	}
	s.manualUnreadMu.RLock()
	unread, ok := s.manualUnread[key]
	s.manualUnreadMu.RUnlock()
	return unread, ok
}

func (s *AppState) setCurrentChatMessages(messages []csa.ChatMessage) {
	copied := append([]csa.ChatMessage(nil), messages...)
	s.chatMessagesMu.Lock()
	s.chatMessages = copied
	s.chatRowMap = nil
	s.chatMessagesMu.Unlock()
}

func (s *AppState) setCurrentChatRowMap(rowMap []int) {
	copied := append([]int(nil), rowMap...)
	s.chatMessagesMu.Lock()
	s.chatRowMap = copied
	s.chatMessagesMu.Unlock()
}

func (s *AppState) getCurrentChatMessage(index int) (csa.ChatMessage, bool) {
	s.chatMessagesMu.RLock()
	defer s.chatMessagesMu.RUnlock()
	if index < 0 {
		return csa.ChatMessage{}, false
	}
	messageIdx := index
	if len(s.chatRowMap) > 0 {
		if index >= len(s.chatRowMap) {
			return csa.ChatMessage{}, false
		}
		messageIdx = s.chatRowMap[index]
	}
	if messageIdx < 0 || messageIdx >= len(s.chatMessages) {
		return csa.ChatMessage{}, false
	}
	return s.chatMessages[messageIdx], true
}

func summarizeReplyPreview(content string) string {
	preview := strings.TrimSpace(strings.Join(strings.Fields(textMessage(content)), " "))
	if preview == "" {
		return "(message)"
	}
	if len(preview) > 80 {
		return preview[:77] + "..."
	}
	return preview
}

func reactionMessageKey(conversationID, messageID string) string {
	conv := normalizeFavoriteKey(conversationID)
	msg := normalizeFavoriteKey(messageID)
	if conv == "" || msg == "" {
		return ""
	}
	return conv + "|" + msg
}

func (s *AppState) setLocalMessageReaction(conversationID, messageID, reaction string) {
	key := reactionMessageKey(conversationID, messageID)
	if key == "" {
		return
	}
	s.messageReactionsMu.Lock()
	if s.messageReactions == nil {
		s.messageReactions = map[string]string{}
	}
	s.messageReactions[key] = strings.TrimSpace(strings.ToLower(reaction))
	s.messageReactionsMu.Unlock()
}

func (s *AppState) getLocalMessageReaction(conversationID, messageID string) string {
	key := reactionMessageKey(conversationID, messageID)
	if key == "" {
		return ""
	}
	s.messageReactionsMu.RLock()
	reaction := s.messageReactions[key]
	s.messageReactionsMu.RUnlock()
	return reaction
}

func (s *AppState) formatMessageSecondary(message csa.ChatMessage, author string) string {
	secondary := strings.TrimSpace(author)
	if secondary == "" {
		secondary = "Unknown"
	}
	secondary = fmt.Sprintf("[%s]%s[-]", s.authorStyleTag(), secondary)
	parts := []string{}
	for _, emotion := range message.Properties.Emotions {
		name := strings.TrimSpace(emotion.Key)
		if name == "" {
			continue
		}
		parts = append(parts, fmt.Sprintf("%s(%d)", name, len(emotion.Users)))
	}
	if local := s.getLocalMessageReaction(message.ConversationId, message.Id); local != "" {
		parts = append(parts, "you:"+local)
	}
	if len(parts) > 0 {
		return secondary + " [gray]| Reactions: " + strings.Join(parts, ", ") + "[-]"
	}
	return secondary
}

func (s *AppState) formatMessageReactionsOnly(message csa.ChatMessage) string {
	parts := []string{}
	for _, emotion := range message.Properties.Emotions {
		name := strings.TrimSpace(emotion.Key)
		if name == "" {
			continue
		}
		parts = append(parts, fmt.Sprintf("%s(%d)", name, len(emotion.Users)))
	}
	if local := s.getLocalMessageReaction(message.ConversationId, message.Id); local != "" {
		parts = append(parts, "you:"+local)
	}
	return strings.Join(parts, ", ")
}

func (s *AppState) setPendingReply(reply *replyTarget) {
	s.replyMu.Lock()
	if reply == nil {
		s.pendingReply = nil
		s.replyMu.Unlock()
		return
	}
	copied := *reply
	s.pendingReply = &copied
	s.replyMu.Unlock()
}

func (s *AppState) getPendingReply() *replyTarget {
	s.replyMu.RLock()
	defer s.replyMu.RUnlock()
	if s.pendingReply == nil {
		return nil
	}
	copied := *s.pendingReply
	return &copied
}

func (s *AppState) clearPendingReply() {
	s.replyMu.Lock()
	s.pendingReply = nil
	s.replyMu.Unlock()
}

func flattenConversationNodes(root *tview.TreeNode) []*tview.TreeNode {
	if root == nil {
		return nil
	}

	var nodes []*tview.TreeNode
	for _, child := range root.GetChildren() {
		if _, ok := child.GetReference().(conversationRef); ok {
			nodes = append(nodes, child)
			continue
		}
		nodes = append(nodes, flattenConversationNodes(child)...)
	}
	return nodes
}

func moveChatNodeToGroup(chatsNode, favoritesNode, recentNode, node *tview.TreeNode, favorite bool) {
	if node == nil || chatsNode == nil || favoritesNode == nil || recentNode == nil {
		return
	}
	favoritesNode.RemoveChild(node)
	recentNode.RemoveChild(node)

	if favorite {
		favoritesNode.AddChild(node)
	} else {
		recentNode.AddChild(node)
	}

	chatsNode.ClearChildren()
	if len(favoritesNode.GetChildren()) > 0 {
		chatsNode.AddChild(favoritesNode)
	}
	if len(recentNode.GetChildren()) > 0 {
		chatsNode.AddChild(recentNode)
	}
}

func isPrivateNotesChat(chat csa.Chat) bool {
	return isPrivateNotesConversationID(chat.Id)
}

func isPrivateNotesConversationID(conversationID string) bool {
	return strings.Contains(strings.ToLower(strings.TrimSpace(conversationID)), "48:notes")
}

func chatFavoriteKey(chatID string, conversationIDs []string) string {
	if key := normalizeFavoriteKey(chatID); key != "" {
		return key
	}
	for _, id := range conversationIDs {
		if key := normalizeFavoriteKey(id); key != "" {
			return key
		}
	}
	return ""
}

func normalizeFavoriteKey(chatID string) string {
	id := strings.TrimSpace(strings.ToLower(chatID))
	if id == "" {
		return ""
	}
	if idx := strings.Index(id, "?"); idx >= 0 {
		id = id[:idx]
	}
	if idx := strings.Index(id, "#"); idx >= 0 {
		id = id[:idx]
	}
	return strings.TrimSpace(id)
}

func ensurePrivateNotesChat(chats []csa.Chat, feeds []csa.PrivateFeed) []csa.Chat {
	for _, chat := range chats {
		if isPrivateNotesChat(chat) {
			return chats
		}
	}

	for _, feed := range feeds {
		notesID := extractConversationID(feed.TargetLink)
		if notesID == "" && isPrivateNotesConversationID(feed.Id) {
			notesID = feed.Id
		}
		if !isPrivateNotesConversationID(notesID) {
			continue
		}
		return append(chats, csa.Chat{
			Id:          notesID,
			Title:       "Private Notes",
			LastMessage: feed.LastMessage,
		})
	}

	return chats
}

func extractConversationID(targetLink string) string {
	link := strings.TrimSpace(targetLink)
	if link == "" {
		return ""
	}
	lower := strings.ToLower(link)

	if start := strings.Index(lower, "/chat/"); start >= 0 {
		start += len("/chat/")
		rest := link[start:]
		restLower := lower[start:]
		if end := strings.Index(restLower, "/conversations"); end >= 0 {
			return normalizeFavoriteKey(rest[:end])
		}
	}

	if start := strings.Index(lower, "/conversations/"); start >= 0 {
		start += len("/conversations/")
		rest := link[start:]
		end := len(rest)
		for _, marker := range []string{"/", "?", "#"} {
			if idx := strings.Index(rest, marker); idx >= 0 && idx < end {
				end = idx
			}
		}
		return normalizeFavoriteKey(rest[:end])
	}

	return ""
}

func textMessage(input string) string {
	output := ""
	z := html.NewTokenizer(bytes.NewBuffer([]byte(input)))
	for {
		tt := z.Next()
		if tt == html.ErrorToken {
			break
		}

		switch tt {
		case html.TextToken:
			text := string(z.Text())
			if strings.TrimSpace(text) == "" {
				continue
			}
			output += fmt.Sprintf("%v\n", text)
		}
		if tt == html.ErrorToken {
			break
		}
	}
	return output
}

func (s *AppState) loadConversations(c *csa.Channel) {
	s.loadConversationsByIDs(nil, []string{c.Id}, c.DisplayName)
}

func (s *AppState) sendMessageAndRefresh(conversationIDs []string, displayName, content string, selectedNode *tview.TreeNode, reply *replyTarget) {
	err := s.sendMessage(conversationIDs, content, reply)
	if err != nil {
		s.showError(err)
		return
	}
	// The send endpoint may acknowledge before the message is visible in reads.
	for _, delay := range []time.Duration{0, 300 * time.Millisecond, 1 * time.Second, 2 * time.Second} {
		if delay > 0 {
			time.Sleep(delay)
		}
		s.loadConversationsByIDs(selectedNode, conversationIDs, displayName)
	}
}

func formatOutgoingHTML(content string, reply *replyTarget) string {
	lines := strings.Split(content, "\n")
	for i := range lines {
		lines[i] = strings.TrimSpace(lines[i])
	}
	body := strings.Join(lines, "<br/>")
	if reply != nil {
		author := strings.TrimSpace(reply.Author)
		if author == "" {
			author = "message"
		}
		preview := strings.TrimSpace(reply.Preview)
		if preview == "" {
			preview = "(message)"
		}
		quoted := "<blockquote><strong>Reply to " + author + ":</strong><br/>" + preview + "</blockquote>"
		return quoted + "<div><div>" + body + "</div></div>"
	}
	return "<div><div>" + body + "</div></div>"
}

func (s *AppState) sendMessage(conversationIDs []string, content string, reply *replyTarget) error {
	ids := normalizeConversationIDs(conversationIDs)
	if len(ids) == 0 {
		return fmt.Errorf("no conversation id available")
	}
	mentionContent, mentions := s.applyMentions(content, ids)
	properties := map[string]interface{}{}
	if len(mentions) > 0 {
		mentionsJSON, err := json.Marshal(mentions)
		if err == nil {
			properties["mentions"] = string(mentionsJSON)
		}
	}

	payload := map[string]interface{}{
		"content":         formatOutgoingHTML(mentionContent, reply),
		"messagetype":     "RichText/Html",
		"contenttype":     "text",
		"clientmessageid": strconv.FormatInt(time.Now().UnixNano(), 10),
		"amsreferences":   []string{},
		"properties":      properties,
	}
	bodyBytes, err := json.Marshal(payload)
	if err != nil {
		return fmt.Errorf("unable to encode outgoing message: %v", err)
	}

	var lastErr error
	for _, id := range ids {
		for attempt := 0; attempt < 2; attempt++ {
			endpoint := csa.MessagesHost + "v1/users/ME/conversations/" + url.QueryEscape(id) + "/messages"
			req, err := s.teamsClient.ChatSvc().AuthenticatedRequest(http.MethodPost, endpoint, bytes.NewReader(bodyBytes))
			if err != nil {
				if attempt == 0 && isUnauthorizedError(err) {
					if refreshErr := s.refreshAuthFromTeamsToken(); refreshErr == nil {
						continue
					} else {
						lastErr = fmt.Errorf("unauthorized and unable to refresh auth: %v", refreshErr)
						break
					}
				}
				lastErr = err
				break
			}
			req.Header.Add("Content-Type", "application/json")

			resp, err := http.DefaultClient.Do(req)
			if err != nil {
				lastErr = err
				break
			}

			if resp.StatusCode == http.StatusOK || resp.StatusCode == http.StatusCreated {
				_, _ = io.ReadAll(resp.Body)
				_ = resp.Body.Close()
				s.logger.WithFields(logrus.Fields{
					"conversation_id": id,
					"status_code":     resp.StatusCode,
				}).Info("message sent")
				return nil
			}
			if resp.StatusCode == http.StatusUnauthorized && attempt == 0 {
				_, _ = io.ReadAll(resp.Body)
				_ = resp.Body.Close()
				if refreshErr := s.refreshAuthFromTeamsToken(); refreshErr == nil {
					continue
				} else {
					lastErr = fmt.Errorf("send unauthorized for %s and refresh failed: %v", id, refreshErr)
					break
				}
			}

			body, _ := io.ReadAll(resp.Body)
			_ = resp.Body.Close()
			lastErr = fmt.Errorf("send failed for %s: status=%d body=%s", id, resp.StatusCode, strings.TrimSpace(string(body)))
			s.logger.WithFields(logrus.Fields{
				"conversation_id": id,
				"status_code":     resp.StatusCode,
				"response_body":   strings.TrimSpace(string(body)),
			}).Warn("message send failed for conversation id")
			break
		}
	}

	if lastErr == nil {
		return fmt.Errorf("unable to send message")
	}
	return lastErr
}

func (s *AppState) reactToMessage(message csa.ChatMessage, reaction string) {
	reaction = strings.TrimSpace(strings.ToLower(reaction))
	if reaction == "" {
		reaction = defaultReactionKey
	}
	conversationID := strings.TrimSpace(message.ConversationId)
	messageID := strings.TrimSpace(message.Id)
	if conversationID == "" || messageID == "" {
		return
	}

	s.setLocalMessageReaction(conversationID, messageID, reaction)
	if err := s.sendReaction(message, reaction); err != nil {
		s.logger.WithError(err).WithFields(logrus.Fields{
			"conversation_id": conversationID,
			"message_id":      messageID,
			"reaction":        reaction,
		}).Warn("unable to send reaction to server; keeping local reaction")
	}

	ids, title, node := s.getActiveConversation()
	if len(ids) > 0 {
		s.loadConversationsByIDs(node, ids, title)
	}
}

func (s *AppState) sendReaction(message csa.ChatMessage, reaction string) error {
	conversationID := strings.TrimSpace(message.ConversationId)
	messageID := strings.TrimSpace(message.Id)
	if conversationID == "" || messageID == "" {
		return fmt.Errorf("missing conversation or message id for reaction")
	}

	userMri := ""
	if s.me != nil {
		userMri = strings.TrimSpace(s.me.Mri)
		if userMri == "" && strings.TrimSpace(s.me.ObjectId) != "" {
			userMri = "8:orgid:" + strings.TrimSpace(s.me.ObjectId)
		}
	}
	emotionsPayload, err := buildEmotionsPayload(message, reaction, userMri)
	if err != nil {
		return fmt.Errorf("unable to encode reaction payload: %v", err)
	}

	bodyMapEmotions, err := json.Marshal(map[string]interface{}{
		"emotions": emotionsPayload,
	})
	if err != nil {
		return fmt.Errorf("unable to encode reaction emotions map: %v", err)
	}
	bodyMapProperties, err := json.Marshal(map[string]interface{}{
		"properties": map[string]interface{}{
			"emotions": emotionsPayload,
		},
	})
	if err != nil {
		return fmt.Errorf("unable to encode reaction properties map: %v", err)
	}
	bodyValueOnly, err := json.Marshal(map[string]interface{}{
		"value": emotionsPayload,
	})
	if err != nil {
		return fmt.Errorf("unable to encode reaction value payload: %v", err)
	}
	bodyJSONString, err := json.Marshal(emotionsPayload)
	if err != nil {
		return fmt.Errorf("unable to encode reaction json string payload: %v", err)
	}

	base := csa.MessagesHost + "v1/users/ME/conversations/" + url.QueryEscape(conversationID) + "/messages/" + url.QueryEscape(messageID)
	type reactionAttempt struct {
		method   string
		endpoint string
		body     []byte
	}
	attempts := []reactionAttempt{
		{method: http.MethodPut, endpoint: base + "/properties", body: bodyMapEmotions},
		{method: http.MethodPatch, endpoint: base + "/properties", body: bodyMapEmotions},
		{method: http.MethodPatch, endpoint: base, body: bodyMapProperties},
		{method: http.MethodPut, endpoint: base + "/properties?name=emotions", body: bodyJSONString},
		{method: http.MethodPut, endpoint: base + "/properties?name=emotions", body: bodyValueOnly},
		{method: http.MethodPut, endpoint: base + "/properties?name=emotions", body: bodyMapEmotions},
	}
	var lastErr error
	for _, attempt := range attempts {
		err = s.sendReactionRequest(attempt.method, attempt.endpoint, attempt.body)
		if err == nil {
			s.logger.WithFields(logrus.Fields{
				"method":          attempt.method,
				"endpoint":        attempt.endpoint,
				"conversation_id": conversationID,
				"message_id":      messageID,
			}).Info("reaction sent")
			return nil
		}
		lastErr = err
	}
	if lastErr == nil {
		lastErr = fmt.Errorf("reaction request failed")
	}
	return lastErr
}

func buildEmotionsPayload(message csa.ChatMessage, reaction, userMri string) (string, error) {
	type userEmotionWire struct {
		Mri   string `json:"mri"`
		Time  int64  `json:"time"`
		Value string `json:"value"`
	}
	type emotionWire struct {
		Key   string            `json:"key"`
		Users []userEmotionWire `json:"users"`
	}

	byKey := map[string]map[string]userEmotionWire{}
	order := []string{}
	for _, emotion := range message.Properties.Emotions {
		key := strings.TrimSpace(strings.ToLower(emotion.Key))
		if key == "" {
			continue
		}
		if _, ok := byKey[key]; !ok {
			byKey[key] = map[string]userEmotionWire{}
			order = append(order, key)
		}
		for _, u := range emotion.Users {
			mri := strings.TrimSpace(u.Mri)
			if mri == "" {
				continue
			}
			byKey[key][mri] = userEmotionWire{Mri: mri, Time: u.Time, Value: strings.TrimSpace(u.Value)}
		}
	}

	reaction = strings.TrimSpace(strings.ToLower(reaction))
	if reaction == "" {
		reaction = defaultReactionKey
	}
	if _, ok := byKey[reaction]; !ok {
		byKey[reaction] = map[string]userEmotionWire{}
		order = append(order, reaction)
	}
	if strings.TrimSpace(userMri) != "" {
		byKey[reaction][userMri] = userEmotionWire{
			Mri:   strings.TrimSpace(userMri),
			Time:  time.Now().UnixMilli(),
			Value: reaction,
		}
	}

	wires := []emotionWire{}
	for _, key := range order {
		usersMap := byKey[key]
		users := make([]userEmotionWire, 0, len(usersMap))
		for _, u := range usersMap {
			users = append(users, u)
		}
		sort.Slice(users, func(i, j int) bool {
			return users[i].Mri < users[j].Mri
		})
		wires = append(wires, emotionWire{Key: key, Users: users})
	}
	bytesOut, err := json.Marshal(wires)
	if err != nil {
		return "", err
	}
	return string(bytesOut), nil
}

func (s *AppState) sendReactionRequest(method, endpoint string, body []byte) error {
	var lastErr error
	for attempt := 0; attempt < 2; attempt++ {
		req, err := s.teamsClient.ChatSvc().AuthenticatedRequest(method, endpoint, bytes.NewReader(body))
		if err != nil {
			if attempt == 0 && isUnauthorizedError(err) {
				if refreshErr := s.refreshAuthFromTeamsToken(); refreshErr == nil {
					continue
				}
			}
			return err
		}
		req.Header.Set("Content-Type", "application/json")

		resp, err := http.DefaultClient.Do(req)
		if err != nil {
			lastErr = err
			break
		}
		respBody, _ := io.ReadAll(resp.Body)
		_ = resp.Body.Close()

		if resp.StatusCode == http.StatusOK || resp.StatusCode == http.StatusCreated || resp.StatusCode == http.StatusNoContent {
			return nil
		}
		if resp.StatusCode == http.StatusUnauthorized && attempt == 0 {
			if refreshErr := s.refreshAuthFromTeamsToken(); refreshErr == nil {
				continue
			}
		}
		lastErr = fmt.Errorf("%s %s status=%d body=%s", method, endpoint, resp.StatusCode, strings.TrimSpace(string(respBody)))
		break
	}
	if lastErr == nil {
		lastErr = fmt.Errorf("reaction request failed")
	}
	return lastErr
}

func normalizeConversationIDs(conversationIDs []string) []string {
	ids := []string{}
	seen := map[string]struct{}{}
	for _, id := range conversationIDs {
		id = strings.TrimSpace(id)
		if id == "" {
			continue
		}
		if _, ok := seen[id]; ok {
			continue
		}
		seen[id] = struct{}{}
		ids = append(ids, id)
	}
	return ids
}

func extractTrailingMentionQuery(text string) (start int, prefix string, query string, ok bool) {
	if strings.TrimSpace(text) == "" {
		return 0, "", "", false
	}
	if strings.HasSuffix(text, "c@") || strings.HasSuffix(text, "C@") {
		return len(text) - 2, "c@", "", true
	}
	if strings.HasSuffix(text, "@") {
		return len(text) - 1, "@", "", true
	}
	end := len(text)
	i := end - 1
	for i >= 0 {
		ch := text[i]
		if (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9') || ch == '.' || ch == '_' || ch == '-' {
			i--
			continue
		}
		break
	}
	tokenStart := i + 1
	if tokenStart <= 0 {
		return 0, "", "", false
	}
	if tokenStart >= len(text) {
		return 0, "", "", false
	}
	if text[tokenStart-1] == '@' {
		if tokenStart-2 >= 0 && (text[tokenStart-2] == 'c' || text[tokenStart-2] == 'C') {
			return tokenStart - 2, "c@", text[tokenStart:end], true
		}
		return tokenStart - 1, "@", text[tokenStart:end], true
	}
	return 0, "", "", false
}

func mentionTokenFromDisplayName(name string) string {
	name = strings.TrimSpace(name)
	if name == "" {
		return ""
	}
	var b strings.Builder
	for _, r := range name {
		if (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') || (r >= '0' && r <= '9') || r == '.' || r == '_' || r == '-' {
			b.WriteRune(r)
		}
	}
	return strings.TrimSpace(b.String())
}

func normalizeMentionSearch(s string) string {
	s = strings.ToLower(strings.TrimSpace(s))
	if s == "" {
		return ""
	}
	var b strings.Builder
	for _, r := range s {
		if (r >= 'a' && r <= 'z') || (r >= '0' && r <= '9') {
			b.WriteRune(r)
		}
	}
	return b.String()
}

func (s *AppState) findMentionSuggestions(prefix, query string, conversationIDs []string) []mentionCandidate {
	var base []mentionCandidate
	if strings.EqualFold(prefix, "c@") {
		base = s.mentionCandidatesGlobal()
	} else {
		base = append([]mentionCandidate{}, s.mentionCandidatesForConversation(conversationIDs)...)
		base = append(base, s.mentionCandidatesGlobal()...)
		base = uniqueMentions(base)
	}
	queryNorm := normalizeMentionSearch(query)
	if queryNorm == "" {
		return base
	}
	starts := []mentionCandidate{}
	contains := []mentionCandidate{}
	for _, c := range base {
		nameNorm := normalizeMentionSearch(c.DisplayName)
		if strings.HasPrefix(nameNorm, queryNorm) {
			starts = append(starts, c)
		} else if strings.Contains(nameNorm, queryNorm) {
			contains = append(contains, c)
		}
	}
	return append(starts, contains...)
}

func inferMriFromMessageFrom(from string) string {
	from = strings.TrimSpace(from)
	if from == "" {
		return ""
	}
	if strings.Contains(from, "/contacts/") {
		parts := strings.Split(from, "/contacts/")
		if len(parts) > 1 {
			return strings.TrimSpace(parts[len(parts)-1])
		}
	}
	return ""
}

func (s *AppState) mentionCandidatesFromCurrentMessages() []mentionCandidate {
	s.chatMessagesMu.RLock()
	msgs := append([]csa.ChatMessage(nil), s.chatMessages...)
	s.chatMessagesMu.RUnlock()
	candidates := []mentionCandidate{}
	for _, m := range msgs {
		name := strings.TrimSpace(m.ImDisplayName)
		if name == "" {
			name = inferMessageAuthor(m, s.me)
		}
		if isSelfDisplayName(name, s.me) || strings.EqualFold(strings.TrimSpace(name), "you") || strings.EqualFold(strings.TrimSpace(name), "unknown") {
			continue
		}
		mri := inferMriFromMessageFrom(m.From)
		objectID := ""
		if strings.HasPrefix(strings.ToLower(mri), "8:orgid:") {
			objectID = strings.TrimPrefix(strings.TrimPrefix(strings.ToLower(mri), "8:orgid:"), " ")
		}
		candidates = append(candidates, mentionCandidate{
			DisplayName: name,
			Mri:         mri,
			ObjectID:    objectID,
		})
	}
	return uniqueMentions(candidates)
}

func (s *AppState) resetMentionCycle() {
	s.mentionCycleMu.Lock()
	s.mentionCycleToken = ""
	s.mentionCycleIndex = -1
	s.mentionCycleItems = nil
	s.mentionCycleMu.Unlock()
}

func (s *AppState) getMentionCycleSuggestions(prefix, query, tokenKey, text string, conversationIDs []string) ([]mentionCandidate, string) {
	s.mentionCycleMu.Lock()
	defer s.mentionCycleMu.Unlock()

	if len(s.mentionCycleItems) > 0 && strings.HasPrefix(s.mentionCycleToken, strings.ToLower(prefix+":")) {
		for _, c := range s.mentionCycleItems {
			tok := prefix + mentionTokenFromDisplayName(c.DisplayName)
			if strings.HasSuffix(text, tok) {
				copied := append([]mentionCandidate(nil), s.mentionCycleItems...)
				return copied, s.mentionCycleToken
			}
		}
	}

	suggestions := s.findMentionSuggestions(prefix, query, conversationIDs)
	s.mentionCycleToken = tokenKey
	s.mentionCycleIndex = -1
	s.mentionCycleItems = append([]mentionCandidate(nil), suggestions...)
	return suggestions, tokenKey
}

func (s *AppState) setMentionCycleItems(tokenKey string, items []mentionCandidate) {
	s.mentionCycleMu.Lock()
	s.mentionCycleToken = tokenKey
	s.mentionCycleItems = append([]mentionCandidate(nil), items...)
	s.mentionCycleMu.Unlock()
}

func (s *AppState) nextMentionCycleIndex(token string, count int, backwards bool) int {
	if count <= 0 {
		return -1
	}
	s.mentionCycleMu.Lock()
	defer s.mentionCycleMu.Unlock()
	if s.mentionCycleToken != token {
		s.mentionCycleToken = token
		s.mentionCycleIndex = -1
	}
	if backwards {
		if s.mentionCycleIndex <= 0 {
			s.mentionCycleIndex = count - 1
		} else {
			s.mentionCycleIndex--
		}
	} else {
		s.mentionCycleIndex = (s.mentionCycleIndex + 1) % count
	}
	return s.mentionCycleIndex
}

func (s *AppState) applyMentions(content string, conversationIDs []string) (string, []mentionWire) {
	if strings.TrimSpace(content) == "" {
		return content, nil
	}
	conversationCandidates := s.mentionCandidatesForConversation(conversationIDs)
	globalCandidates := s.mentionCandidatesGlobal()
	mentions := []mentionWire{}
	seenKeyToID := map[string]int{}

	out := mentionTokenRegex.ReplaceAllStringFunc(content, func(match string) string {
		prefixWhitespace := ""
		if len(match) > 0 {
			first := match[0]
			if first == ' ' || first == '\t' || first == '\n' {
				prefixWhitespace = string(first)
				match = match[1:]
			}
		}
		atPrefix := "@"
		query := ""
		if strings.HasPrefix(strings.ToLower(match), "c@") {
			atPrefix = "c@"
			query = match[2:]
		} else if strings.HasPrefix(match, "@") {
			atPrefix = "@"
			query = match[1:]
		} else {
			return prefixWhitespace + match
		}
		if strings.TrimSpace(query) == "" {
			return prefixWhitespace + match
		}

		var candidate mentionCandidate
		var ok bool
		if atPrefix == "@" {
			candidate, ok = findMentionCandidate(query, conversationCandidates)
			if !ok {
				candidate, ok = findMentionCandidate(query, globalCandidates)
			}
		} else {
			candidate, ok = findMentionCandidate(query, globalCandidates)
		}
		if !ok || strings.TrimSpace(candidate.DisplayName) == "" {
			return prefixWhitespace + match
		}
		mri := strings.TrimSpace(candidate.Mri)
		if mri == "" && strings.TrimSpace(candidate.ObjectID) != "" {
			mri = "8:orgid:" + strings.TrimSpace(candidate.ObjectID)
		}

		key := strings.ToLower(strings.TrimSpace(candidate.DisplayName)) + "|" + strings.ToLower(strings.TrimSpace(mri))
		mentionID, exists := seenKeyToID[key]
		if !exists {
			mentionID = len(mentions)
			seenKeyToID[key] = mentionID
			mentions = append(mentions, mentionWire{
				ID:          mentionID,
				MentionType: "person",
				Mri:         mri,
				DisplayName: strings.TrimSpace(candidate.DisplayName),
				ObjectId:    strings.TrimSpace(candidate.ObjectID),
			})
		}
		return prefixWhitespace + fmt.Sprintf("<at id=\"%d\">@%s</at>", mentionID, candidate.DisplayName)
	})

	return out, mentions
}

func findMentionCandidate(query string, candidates []mentionCandidate) (mentionCandidate, bool) {
	query = normalizeMentionSearch(query)
	if query == "" {
		return mentionCandidate{}, false
	}
	for _, c := range candidates {
		name := normalizeMentionSearch(c.DisplayName)
		if strings.HasPrefix(name, query) {
			return c, true
		}
	}
	for _, c := range candidates {
		name := normalizeMentionSearch(c.DisplayName)
		if strings.Contains(name, query) {
			return c, true
		}
	}
	return mentionCandidate{}, false
}

func mentionKey(c mentionCandidate) string {
	name := strings.ToLower(strings.TrimSpace(c.DisplayName))
	mri := strings.ToLower(strings.TrimSpace(c.Mri))
	obj := strings.ToLower(strings.TrimSpace(c.ObjectID))
	if mri == "" && obj != "" {
		mri = "8:orgid:" + obj
	}
	if name == "" {
		return ""
	}
	return name + "|" + mri + "|" + obj
}

func uniqueMentions(in []mentionCandidate) []mentionCandidate {
	out := make([]mentionCandidate, 0, len(in))
	seen := map[string]struct{}{}
	for _, c := range in {
		if strings.TrimSpace(c.DisplayName) == "" {
			continue
		}
		key := mentionKey(c)
		if key == "" {
			continue
		}
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = struct{}{}
		out = append(out, c)
	}
	sort.Slice(out, func(i, j int) bool {
		return strings.ToLower(out[i].DisplayName) < strings.ToLower(out[j].DisplayName)
	})
	return out
}

func (s *AppState) mentionCandidatesForConversation(conversationIDs []string) []mentionCandidate {
	if s.conversations == nil {
		return s.mentionCandidatesFromCurrentMessages()
	}
	ids := map[string]struct{}{}
	for _, id := range conversationIDs {
		n := normalizeFavoriteKey(id)
		if n != "" {
			ids[n] = struct{}{}
		}
	}
	if len(ids) == 0 {
		return nil
	}

	candidates := []mentionCandidate{}
	candidates = append(candidates, s.mentionCandidatesFromCurrentMessages()...)
	for _, chat := range s.conversations.Chats {
		matched := false
		for _, cid := range candidateConversationIds(chat, s.conversations.PrivateFeeds) {
			if _, ok := ids[normalizeFavoriteKey(cid)]; ok {
				matched = true
				break
			}
		}
		if !matched {
			continue
		}
		for _, member := range chat.Members {
			if isCurrentUser(member, s.me) {
				continue
			}
			name := strings.TrimSpace(member.FriendlyName)
			if isSelfDisplayName(name, s.me) {
				continue
			}
			candidates = append(candidates, mentionCandidate{
				DisplayName: name,
				Mri:         strings.TrimSpace(member.Mri),
				ObjectID:    strings.TrimSpace(member.ObjectId),
			})
		}
	}
	return uniqueMentions(candidates)
}

func (s *AppState) mentionCandidatesGlobal() []mentionCandidate {
	candidates := []mentionCandidate{}
	candidates = append(candidates, s.getContactCandidates()...)
	candidates = append(candidates, s.mentionCandidatesFromCurrentMessages()...)
	if s.conversations == nil {
		return uniqueMentions(candidates)
	}
	for _, chat := range s.conversations.Chats {
		for _, member := range chat.Members {
			if isCurrentUser(member, s.me) {
				continue
			}
			name := strings.TrimSpace(member.FriendlyName)
			if isSelfDisplayName(name, s.me) {
				continue
			}
			candidates = append(candidates, mentionCandidate{
				DisplayName: name,
				Mri:         strings.TrimSpace(member.Mri),
				ObjectID:    strings.TrimSpace(member.ObjectId),
			})
		}
	}
	return uniqueMentions(candidates)
}

func (s *AppState) getContactCandidates() []mentionCandidate {
	s.contactsMu.RLock()
	if time.Since(s.contactsLastFetched) < 10*time.Minute && len(s.contactCandidates) > 0 {
		cached := append([]mentionCandidate(nil), s.contactCandidates...)
		s.contactsMu.RUnlock()
		return cached
	}
	s.contactsMu.RUnlock()

	fresh, err := s.fetchContactCandidatesFromAPI()
	if err != nil {
		s.logger.WithError(err).Debug("contacts endpoint unavailable; using fallback participants")
		return nil
	}
	fresh = uniqueMentions(fresh)
	s.contactsMu.Lock()
	s.contactCandidates = append([]mentionCandidate(nil), fresh...)
	s.contactsLastFetched = time.Now()
	s.contactsMu.Unlock()
	return fresh
}

func (s *AppState) fetchContactCandidatesFromAPI() ([]mentionCandidate, error) {
	endpoints := []string{
		csa.MessagesHost + "v1/users/ME/contacts",
		csa.MessagesHost + "v1/users/ME/people",
		"https://teams.microsoft.com/api/mt/part/emea-02/beta/users/people",
	}
	var lastErr error
	for _, ep := range endpoints {
		req, err := s.teamsClient.ChatSvc().AuthenticatedRequest(http.MethodGet, ep, nil)
		if err != nil {
			lastErr = err
			continue
		}
		resp, err := http.DefaultClient.Do(req)
		if err != nil {
			lastErr = err
			continue
		}
		body, _ := io.ReadAll(resp.Body)
		_ = resp.Body.Close()
		if resp.StatusCode < 200 || resp.StatusCode >= 300 {
			lastErr = fmt.Errorf("contacts endpoint %s status=%d", ep, resp.StatusCode)
			continue
		}
		candidates := extractMentionCandidatesFromJSON(body)
		if len(candidates) > 0 {
			return candidates, nil
		}
		lastErr = fmt.Errorf("contacts endpoint %s returned no candidates", ep)
	}
	if lastErr == nil {
		lastErr = fmt.Errorf("no contacts endpoint succeeded")
	}
	return nil, lastErr
}

func extractMentionCandidatesFromJSON(data []byte) []mentionCandidate {
	var raw interface{}
	if err := json.Unmarshal(data, &raw); err != nil {
		return nil
	}
	out := []mentionCandidate{}
	var walk func(v interface{})
	walk = func(v interface{}) {
		switch t := v.(type) {
		case map[string]interface{}:
			name := firstString(t, "displayName", "friendlyName", "name")
			mri := firstString(t, "mri", "id")
			objectID := firstString(t, "objectId", "oid")
			if strings.TrimSpace(name) != "" {
				out = append(out, mentionCandidate{
					DisplayName: strings.TrimSpace(name),
					Mri:         strings.TrimSpace(mri),
					ObjectID:    strings.TrimSpace(objectID),
				})
			}
			for _, v2 := range t {
				walk(v2)
			}
		case []interface{}:
			for _, v2 := range t {
				walk(v2)
			}
		}
	}
	walk(raw)
	return out
}

func firstString(m map[string]interface{}, keys ...string) string {
	for _, k := range keys {
		if v, ok := m[k]; ok {
			if s, ok := v.(string); ok {
				s = strings.TrimSpace(s)
				if s != "" {
					return s
				}
			}
		}
	}
	return ""
}

func (s *AppState) fetchConversationMessages(displayName string, conversationIDs []string) ([]string, []csa.ChatMessage, error) {
	ids := normalizeConversationIDs(conversationIDs)
	if len(ids) == 0 {
		return nil, nil, fmt.Errorf("no conversation id available")
	}

	var messages []csa.ChatMessage
	var err error
	for idx, id := range ids {
		s.logger.WithFields(logrus.Fields{
			"display_name":    displayName,
			"conversation_id": id,
			"attempt":         strconv.Itoa(idx + 1),
		}).Debug("fetching messages")
		for attempt := 0; attempt < 2; attempt++ {
			messages, err = s.teamsClient.GetMessages(&csa.Channel{Id: id, DisplayName: displayName})
			if err == nil {
				s.logger.WithFields(logrus.Fields{
					"display_name":    displayName,
					"conversation_id": id,
					"attempt":         strconv.Itoa(idx + 1),
					"messages_count":  len(messages),
				}).Info("messages loaded")
				break
			}
			if attempt == 0 && isUnauthorizedError(err) {
				refreshErr := s.refreshAuthFromTeamsToken()
				if refreshErr == nil {
					continue
				}
				s.logger.WithFields(logrus.Fields{
					"display_name":    displayName,
					"conversation_id": id,
					"refresh_error":   refreshErr.Error(),
				}).Warn("unable to refresh auth after unauthorized response")
			}
			break
		}
		if err == nil {
			break
		}
		s.logger.WithFields(logrus.Fields{
			"display_name":    displayName,
			"conversation_id": id,
			"attempt":         strconv.Itoa(idx + 1),
			"error":           err.Error(),
		}).Warn("message fetch failed for conversation id")
	}
	if err != nil {
		return ids, nil, err
	}

	sort.Sort(csa.SortMessageByTime(messages))
	return ids, messages, nil
}

func (s *AppState) resolveConversationTitle(displayName string, conversationIDs []string) (string, error) {
	_, messages, err := s.fetchConversationMessages(displayName, conversationIDs)
	if err != nil {
		return displayName, err
	}
	return resolveDMDisplayName(displayName, messages, s.me), nil
}

func (s *AppState) loadConversationsByIDs(selectedNode *tview.TreeNode, conversationIDs []string, displayName string) {
	s.logger.WithFields(logrus.Fields{
		"display_name":  displayName,
		"incoming_ids":  strings.Join(conversationIDs, ","),
		"incoming_size": len(conversationIDs),
	}).Debug("load conversations called")
	ids, messages, err := s.fetchConversationMessages(displayName, conversationIDs)
	if err != nil {
		s.logger.WithFields(logrus.Fields{
			"display_name":  displayName,
			"attempted_ids": strings.Join(ids, ","),
		}).WithError(err).Error("all conversation id attempts failed")
		s.showError(err)
		time.Sleep(5 * time.Second)
		s.pages.SwitchToPage(PageMain)
		s.app.Draw()
		s.app.SetFocus(s.pages)
		return
	}

	displayName = resolveDMDisplayName(displayName, messages, s.me)
	if selectedNode != nil && strings.TrimSpace(selectedNode.GetText()) != strings.TrimSpace(displayName) {
		if ref, ok := selectedNode.GetReference().(conversationRef); ok {
			ref.isUnread = false
			s.setManualUnread(ref.chatKey, false)
			ref.title = displayName
			selectedNode.SetText(formatChatTreeTitle(ref.title, ref.isUnread))
			selectedNode.SetReference(ref)
			if s.setChatTitle(ref.chatKey, displayName) {
				s.persistEncryptedChatSettings()
			}
		} else {
			selectedNode.SetText(displayName)
		}
	}
	s.components[ViChat].(*tview.List).
		SetTitle(displayName).
		SetBorder(true).
		SetTitleAlign(tview.AlignCenter)

	// Clear chat
	chatList := s.components[ViChat].(*tview.List)
	chatList.Clear()
	chatList.ShowSecondaryText(true)
	chatList.SetSelectedFunc(nil)
	chatList.SetChangedFunc(nil)
	s.setCurrentChatMessages(messages)
	rowMap := []int{}
	s.logger.WithFields(logrus.Fields{
		"display_name":   displayName,
		"messages_count": len(messages),
	}).Debug("rendering messages")
	_, _, listWidth, _ := chatList.GetRect()
	_, _, _, innerWidth := chatList.GetInnerRect()
	if listWidth > 2 {
		if listWidth-2 > innerWidth {
			innerWidth = listWidth - 2
		}
	}
	if innerWidth <= 0 {
		innerWidth = 80
	}
	wrapWidth := s.getChatWrapPercent()
	if wrapWidth > innerWidth {
		wrapWidth = innerWidth
	}
	if wrapWidth < 8 {
		wrapWidth = 8
	}
	s.chatWordWrapMu.Lock()
	s.chatWrapEffective = wrapWidth
	s.chatWordWrapMu.Unlock()
	for msgIdx, message := range messages {
		author := strings.TrimSpace(message.ImDisplayName)
		if author == "" {
			author = inferMessageAuthor(message, s.me)
		}
		if s.isChatWordWrap() {
			lines := wrapTextLines(textMessage(message.Content), wrapWidth)
			if len(lines) == 0 {
				lines = []string{""}
			}
			secondary := s.formatMessageSecondary(message, author)
			for i, line := range lines {
				lineSecondary := ""
				if i == 0 {
					lineSecondary = secondary
				}
				chatList.AddItem(line, lineSecondary, 0, nil)
				rowMap = append(rowMap, msgIdx)
			}
		} else {
			chatList.AddItem(s.formatChatMessageText(message.Content), s.formatMessageSecondary(message, author), 0, nil)
			rowMap = append(rowMap, msgIdx)
		}
	}
	s.setCurrentChatRowMap(rowMap)
	if chatList.GetItemCount() > 0 {
		chatList.SetCurrentItem(chatList.GetItemCount() - 1)
	}
	s.app.Draw()
}

func inferMessageAuthor(message csa.ChatMessage, me *models.User) string {
	if me != nil {
		from := strings.ToLower(strings.TrimSpace(message.From))
		if from != "" {
			if strings.TrimSpace(me.ObjectId) != "" && strings.Contains(from, strings.ToLower(me.ObjectId)) {
				if strings.TrimSpace(me.DisplayName) != "" {
					return me.DisplayName
				}
				return "You"
			}
			if strings.TrimSpace(me.Mri) != "" && strings.Contains(from, strings.ToLower(me.Mri)) {
				if strings.TrimSpace(me.DisplayName) != "" {
					return me.DisplayName
				}
				return "You"
			}
		}
	}
	return "Unknown"
}

func defaultSettingsPaths() (string, string) {
	homeDir, err := os.UserHomeDir()
	if err != nil || strings.TrimSpace(homeDir) == "" {
		return "teams-cli-settings.enc", "teams-cli-settings.key"
	}
	configDir := filepath.Join(homeDir, ".config", "fossteams")
	return filepath.Join(configDir, "teams-cli-settings.enc"), filepath.Join(configDir, "teams-cli-settings.key")
}

func defaultKeybindPath() string {
	homeDir, err := os.UserHomeDir()
	if err != nil || strings.TrimSpace(homeDir) == "" {
		return "teams-cli-keybindings.json"
	}
	return filepath.Join(homeDir, ".config", "fossteams", "teams-cli-keybindings.json")
}

func defaultKeybindingsForPreset(preset string) map[string][]string {
	b := map[string][]string{
		actionToggleScan:     {"m"},
		actionScanNow:        {"shift+m"},
		actionMarkUnread:     {"r"},
		actionToggleFavorite: {"f"},
		actionRefreshTitles:  {"u"},
		actionReloadKeybinds: {"ctrl+r"},
		actionFocusCompose:   {"i"},
		actionReplyMessage:   {"r"},
		actionReactMessage:   {"e"},
		actionMoveDown:       {"down"},
		actionMoveUp:         {"up"},
	}

	switch strings.ToLower(strings.TrimSpace(preset)) {
	case "jk":
		b[actionMoveDown] = append([]string(nil), b[actionMoveDown]...)
		b[actionMoveDown] = append(b[actionMoveDown], "j")
		b[actionMoveUp] = append([]string(nil), b[actionMoveUp]...)
		b[actionMoveUp] = append(b[actionMoveUp], "k", "K")
	case "vim":
		b[actionMoveDown] = append([]string(nil), b[actionMoveDown]...)
		b[actionMoveDown] = append(b[actionMoveDown], "j")
		b[actionMoveUp] = append([]string(nil), b[actionMoveUp]...)
		b[actionMoveUp] = append(b[actionMoveUp], "k", "K")
		b[actionFocusCompose] = []string{"i", "c"}
	case "emacs":
		b[actionMoveDown] = []string{"down", "ctrl+n"}
		b[actionMoveUp] = []string{"up", "ctrl+p"}
		b[actionFocusCompose] = []string{"i", "ctrl+x"}
	}
	return b
}

func mergeKeybindings(base map[string][]string, overrides map[string][]string) map[string][]string {
	out := map[string][]string{}
	for action, keys := range base {
		out[action] = append([]string(nil), keys...)
	}
	for action, keys := range overrides {
		if len(keys) == 0 {
			continue
		}
		out[action] = append([]string(nil), keys...)
	}
	return out
}

func (s *AppState) loadKeybindingsConfig() error {
	if strings.TrimSpace(s.keybindPath) == "" {
		return nil
	}
	data, err := os.ReadFile(s.keybindPath)
	if err != nil {
		if os.IsNotExist(err) {
			return s.writeDefaultKeybindingsConfig()
		}
		return err
	}
	var cfg keybindingConfigFile
	if err := json.Unmarshal(data, &cfg); err != nil {
		return fmt.Errorf("invalid keybinding config json: %v", err)
	}
	preset := strings.ToLower(strings.TrimSpace(cfg.Preset))
	if preset == "" {
		preset = defaultKeybindPreset
	}
	base := defaultKeybindingsForPreset(preset)
	overrides := map[string][]string{}
	for action, keys := range cfg.Bindings {
		if len(keys) == 0 {
			continue
		}
		overrides[action] = append([]string(nil), keys...)
	}
	merged := mergeKeybindings(base, overrides)

	s.keybindMu.Lock()
	s.keybindPreset = preset
	s.keybindings = merged
	s.keybindOverrides = overrides
	s.keybindMu.Unlock()
	return nil
}

func (s *AppState) reloadKeybindingsConfig() error {
	err := s.loadKeybindingsConfig()
	s.keybindParseErr = err
	return err
}

func (s *AppState) saveKeybindingsConfig() error {
	if strings.TrimSpace(s.keybindPath) == "" {
		return nil
	}
	s.keybindMu.RLock()
	cfg := keybindingConfigFile{
		Preset:   s.keybindPreset,
		Bindings: map[string][]string{},
	}
	for action, keys := range s.keybindOverrides {
		if len(keys) == 0 {
			continue
		}
		cfg.Bindings[action] = append([]string(nil), keys...)
	}
	s.keybindMu.RUnlock()

	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}
	if err := os.MkdirAll(filepath.Dir(s.keybindPath), 0o700); err != nil {
		return err
	}
	tmpPath := s.keybindPath + ".tmp"
	if err := os.WriteFile(tmpPath, data, 0o600); err != nil {
		return err
	}
	return os.Rename(tmpPath, s.keybindPath)
}

func (s *AppState) setKeybindPreset(preset string) error {
	preset = strings.ToLower(strings.TrimSpace(preset))
	if preset == "" {
		preset = defaultKeybindPreset
	}
	base := defaultKeybindingsForPreset(preset)
	s.keybindMu.Lock()
	s.keybindPreset = preset
	s.keybindOverrides = map[string][]string{}
	s.keybindings = base
	s.keybindMu.Unlock()
	return s.saveKeybindingsConfig()
}

func (s *AppState) setActionBinding(action, token string) error {
	action = strings.TrimSpace(action)
	token = strings.TrimSpace(token)
	if action == "" || token == "" {
		return fmt.Errorf("action or token is empty")
	}
	s.keybindMu.Lock()
	if s.keybindOverrides == nil {
		s.keybindOverrides = map[string][]string{}
	}
	s.keybindOverrides[action] = []string{token}
	base := defaultKeybindingsForPreset(s.keybindPreset)
	s.keybindings = mergeKeybindings(base, s.keybindOverrides)
	s.keybindMu.Unlock()
	return s.saveKeybindingsConfig()
}

func (s *AppState) resetActionBindingToDefault(action string) error {
	action = strings.TrimSpace(action)
	if action == "" {
		return fmt.Errorf("action is empty")
	}
	s.keybindMu.Lock()
	if s.keybindOverrides != nil {
		delete(s.keybindOverrides, action)
	}
	base := defaultKeybindingsForPreset(s.keybindPreset)
	s.keybindings = mergeKeybindings(base, s.keybindOverrides)
	s.keybindMu.Unlock()
	return s.saveKeybindingsConfig()
}

func (s *AppState) writeDefaultKeybindingsConfig() error {
	if strings.TrimSpace(s.keybindPath) == "" {
		return nil
	}
	cfg := keybindingConfigFile{
		Preset:   defaultKeybindPreset,
		Bindings: map[string][]string{},
	}
	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}
	if err := os.MkdirAll(filepath.Dir(s.keybindPath), 0o700); err != nil {
		return err
	}
	return os.WriteFile(s.keybindPath, data, 0o600)
}

func (s *AppState) bindingMatches(action string, event *tcell.EventKey) bool {
	if event == nil {
		return false
	}
	s.keybindMu.RLock()
	tokens := append([]string(nil), s.keybindings[action]...)
	s.keybindMu.RUnlock()
	for _, token := range tokens {
		if keyTokenMatches(strings.TrimSpace(token), event) {
			return true
		}
	}
	return false
}

func keyTokenMatches(token string, event *tcell.EventKey) bool {
	if token == "" || event == nil {
		return false
	}
	switch strings.ToLower(token) {
	case "up":
		return event.Key() == tcell.KeyUp
	case "down":
		return event.Key() == tcell.KeyDown
	case "left":
		return event.Key() == tcell.KeyLeft
	case "right":
		return event.Key() == tcell.KeyRight
	case "enter":
		return event.Key() == tcell.KeyEnter
	case "esc":
		return event.Key() == tcell.KeyEscape
	case "ctrl+n":
		return event.Key() == tcell.KeyCtrlN
	case "ctrl+p":
		return event.Key() == tcell.KeyCtrlP
	case "ctrl+r":
		return event.Key() == tcell.KeyCtrlR
	case "ctrl+x":
		return event.Key() == tcell.KeyCtrlX
	case "shift+m":
		return event.Key() == tcell.KeyRune && event.Rune() == 'M'
	}
	if len([]rune(token)) == 1 {
		r := []rune(token)[0]
		return event.Key() == tcell.KeyRune && event.Rune() == r
	}
	return false
}

func (s *AppState) loadEncryptedChatSettings() error {
	if strings.TrimSpace(s.settingsPath) == "" || strings.TrimSpace(s.settingsKey) == "" {
		return nil
	}
	data, err := os.ReadFile(s.settingsPath)
	if err != nil {
		if os.IsNotExist(err) {
			return nil
		}
		return err
	}

	var stored encryptedSettingsFile
	if err = json.Unmarshal(data, &stored); err != nil {
		return fmt.Errorf("invalid settings file format: %v", err)
	}
	if stored.Version != 1 {
		return fmt.Errorf("unsupported settings file version: %d", stored.Version)
	}

	key, err := s.readOrCreateSettingsKey()
	if err != nil {
		return err
	}
	nonce, err := base64.StdEncoding.DecodeString(stored.Nonce)
	if err != nil {
		return fmt.Errorf("invalid settings nonce: %v", err)
	}
	ciphertext, err := base64.StdEncoding.DecodeString(stored.Ciphertext)
	if err != nil {
		return fmt.Errorf("invalid settings ciphertext: %v", err)
	}

	block, err := aes.NewCipher(key)
	if err != nil {
		return err
	}
	gcm, err := cipher.NewGCM(block)
	if err != nil {
		return err
	}
	plaintext, err := gcm.Open(nil, nonce, ciphertext, nil)
	if err != nil {
		return fmt.Errorf("unable to decrypt settings: %v", err)
	}

	var settings persistedChatSettings
	if err = json.Unmarshal(plaintext, &settings); err != nil {
		return fmt.Errorf("invalid decrypted settings payload: %v", err)
	}

	s.chatFavoritesMu.Lock()
	if settings.Favorites == nil {
		s.chatFavorites = map[string]bool{}
	} else {
		s.chatFavorites = settings.Favorites
	}
	s.chatFavoritesMu.Unlock()

	s.chatTitlesMu.Lock()
	if settings.Titles == nil {
		s.chatTitles = map[string]string{}
	} else {
		s.chatTitles = settings.Titles
	}
	s.chatTitlesMu.Unlock()

	s.manualUnreadMu.Lock()
	if settings.UnreadOverrides == nil {
		s.manualUnread = map[string]bool{}
	} else {
		s.manualUnread = settings.UnreadOverrides
	}
	s.manualUnreadMu.Unlock()

	s.chatWordWrapMu.Lock()
	if settings.ChatWordWrap == nil {
		s.chatWordWrap = true
	} else {
		s.chatWordWrap = *settings.ChatWordWrap
	}
	if settings.ChatWrapChars != nil && *settings.ChatWrapChars > 0 {
		s.chatWrapChars = *settings.ChatWrapChars
	} else if settings.ChatWrapPercent != nil && *settings.ChatWrapPercent > 0 {
		// Backward compatibility with previous percent-based setting.
		s.chatWrapChars = *settings.ChatWrapPercent
	} else {
		s.chatWrapChars = 80
	}
	s.chatWordWrapMu.Unlock()
	s.themeMu.Lock()
	s.composeColorName = normalizeComposeColorName(settings.ComposeColor)
	s.authorColorName = normalizeAuthorColorName(settings.AuthorColor)
	s.themeMu.Unlock()

	return nil
}

func (s *AppState) persistEncryptedChatSettings() {
	if strings.TrimSpace(s.settingsPath) == "" || strings.TrimSpace(s.settingsKey) == "" {
		return
	}
	key, err := s.readOrCreateSettingsKey()
	if err != nil {
		s.logger.WithError(err).Warn("unable to load settings encryption key")
		return
	}

	settings := persistedChatSettings{
		Favorites:       map[string]bool{},
		Titles:          map[string]string{},
		UnreadOverrides: map[string]bool{},
	}
	s.chatFavoritesMu.RLock()
	for k, v := range s.chatFavorites {
		settings.Favorites[k] = v
	}
	s.chatFavoritesMu.RUnlock()
	s.chatTitlesMu.RLock()
	for k, v := range s.chatTitles {
		if strings.TrimSpace(v) != "" {
			settings.Titles[k] = strings.TrimSpace(v)
		}
	}
	s.chatTitlesMu.RUnlock()
	s.manualUnreadMu.RLock()
	for k, v := range s.manualUnread {
		settings.UnreadOverrides[k] = v
	}
	s.manualUnreadMu.RUnlock()
	wrap := s.isChatWordWrap()
	settings.ChatWordWrap = &wrap
	wrapChars := s.getChatWrapPercent()
	settings.ChatWrapChars = &wrapChars
	s.themeMu.RLock()
	settings.ComposeColor = normalizeComposeColorName(s.composeColorName)
	settings.AuthorColor = normalizeAuthorColorName(s.authorColorName)
	s.themeMu.RUnlock()

	plaintext, err := json.Marshal(settings)
	if err != nil {
		s.logger.WithError(err).Warn("unable to encode chat settings")
		return
	}

	block, err := aes.NewCipher(key)
	if err != nil {
		s.logger.WithError(err).Warn("unable to initialize encryption cipher")
		return
	}
	gcm, err := cipher.NewGCM(block)
	if err != nil {
		s.logger.WithError(err).Warn("unable to initialize encryption mode")
		return
	}
	nonce := make([]byte, gcm.NonceSize())
	if _, err = rand.Read(nonce); err != nil {
		s.logger.WithError(err).Warn("unable to create encryption nonce")
		return
	}
	ciphertext := gcm.Seal(nil, nonce, plaintext, nil)

	encoded, err := json.Marshal(encryptedSettingsFile{
		Version:    1,
		Nonce:      base64.StdEncoding.EncodeToString(nonce),
		Ciphertext: base64.StdEncoding.EncodeToString(ciphertext),
	})
	if err != nil {
		s.logger.WithError(err).Warn("unable to encode encrypted settings payload")
		return
	}

	if err = os.MkdirAll(filepath.Dir(s.settingsPath), 0o700); err != nil {
		s.logger.WithError(err).Warn("unable to create settings directory")
		return
	}
	tmpPath := s.settingsPath + ".tmp"
	if err = os.WriteFile(tmpPath, encoded, 0o600); err != nil {
		s.logger.WithError(err).Warn("unable to write temporary settings file")
		return
	}
	if err = os.Rename(tmpPath, s.settingsPath); err != nil {
		s.logger.WithError(err).Warn("unable to finalize encrypted settings file")
	}
}

func (s *AppState) readOrCreateSettingsKey() ([]byte, error) {
	if strings.TrimSpace(s.settingsKey) == "" {
		return nil, fmt.Errorf("settings key path is not configured")
	}
	key, err := os.ReadFile(s.settingsKey)
	if err == nil {
		if len(key) != 32 {
			return nil, fmt.Errorf("invalid settings key length")
		}
		return key, nil
	}
	if !os.IsNotExist(err) {
		return nil, err
	}

	key = make([]byte, 32)
	if _, err = rand.Read(key); err != nil {
		return nil, err
	}
	if err = os.MkdirAll(filepath.Dir(s.settingsKey), 0o700); err != nil {
		return nil, err
	}
	if err = os.WriteFile(s.settingsKey, key, 0o600); err != nil {
		return nil, err
	}
	return key, nil
}

func isUnauthorizedError(err error) bool {
	if err == nil {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(err.Error()))
	return strings.Contains(lower, "401") || strings.Contains(lower, "unauthorized")
}

func isAuthMissingError(err error) bool {
	if err == nil {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(err.Error()))
	if strings.Contains(lower, "token-") && strings.Contains(lower, ".jwt") {
		return true
	}
	if strings.Contains(lower, "fossteams") && strings.Contains(lower, "no such file") {
		return true
	}
	if strings.Contains(lower, "skypespaces token") && strings.Contains(lower, "no such file") {
		return true
	}
	return false
}

func (s *AppState) refreshAuthFromTeamsToken() error {
	s.authRefreshMu.Lock()
	defer s.authRefreshMu.Unlock()

	deviceErr := error(nil)
	if err := s.refreshAuthFromDeviceCode(); err == nil {
		s.logger.Info("device code auth refresh succeeded")
		newClient, newClientErr := teams_api.New()
		if newClientErr != nil {
			return fmt.Errorf("unable to reinitialize Teams client after token refresh: %v", newClientErr)
		}
		s.teamsClient = newClient
		return nil
	} else {
		deviceErr = err
		s.logger.WithError(err).Warn("device code auth refresh failed, falling back")
	}

	teamsTokenDir, err := findTeamsTokenDir()
	if err != nil {
		return err
	}

	runCmd, runCmdStr, err := buildTeamsTokenCommand(teamsTokenDir)
	if err != nil {
		return err
	}

	cmd := runCmd
	interactive := false
	if wrapped, ok, note, wrapErr := wrapWithTermEverything(runCmdStr, teamsTokenDir); wrapErr != nil {
		return wrapErr
	} else if ok {
		cmd = wrapped
		interactive = true
		if note != "" {
			s.logger.WithField("term_everything", note).Info("auth refresh using term.everything")
		}
	} else if note != "" {
		s.logger.WithField("term_everything", note).Info("auth refresh without term.everything")
	}

	if interactive {
		var runErr error
		s.app.Suspend(func() {
			cmd.Stdin = os.Stdin
			cmd.Stdout = os.Stdout
			cmd.Stderr = os.Stderr
			runErr = cmd.Run()
		})
		if runErr != nil {
			return fmt.Errorf("auth refresh command failed: %v", runErr)
		}
	} else {
		output, runErr := cmd.CombinedOutput()
		if runErr != nil {
			if deviceErr != nil {
				return fmt.Errorf("device code auth failed: %v; electron auth failed: %v (output: %s)", deviceErr, runErr, strings.TrimSpace(string(output)))
			}
			return fmt.Errorf("auth refresh command failed: %v (output: %s)", runErr, strings.TrimSpace(string(output)))
		}
	}
	s.logger.Info("auth refresh command succeeded")

	newClient, newClientErr := teams_api.New()
	if newClientErr != nil {
		return fmt.Errorf("unable to reinitialize Teams client after token refresh: %v", newClientErr)
	}
	s.teamsClient = newClient
	return nil
}

func findTeamsTokenDir() (string, error) {
	candidates := []string{"teams-token", "teams-token-cli"}
	searchRoots := candidateSearchRoots()
	for _, root := range searchRoots {
		for _, candidate := range candidates {
			path := filepath.Join(root, candidate)
			info, err := os.Stat(path)
			if err == nil && info.IsDir() {
				return path, nil
			}
		}
	}
	return "", fmt.Errorf("no teams-token directory found (checked %s)", strings.Join(candidates, ", "))
}

func buildTeamsTokenCommand(teamsTokenDir string) (*exec.Cmd, string, error) {
	binaryPath := filepath.Join(teamsTokenDir, "teams-token")
	if binaryInfo, statErr := os.Stat(binaryPath); statErr == nil && !binaryInfo.IsDir() && binaryInfo.Mode()&0o111 != 0 {
		cmd := exec.Command("./teams-token")
		cmd.Dir = teamsTokenDir
		return cmd, "./teams-token", nil
	}

	if _, statErr := os.Stat(filepath.Join(teamsTokenDir, "go.mod")); statErr == nil {
		cmd := exec.Command("go", "run", ".")
		cmd.Dir = teamsTokenDir
		return cmd, "go run .", nil
	}

	if _, statErr := os.Stat(filepath.Join(teamsTokenDir, "package.json")); statErr == nil {
		if hasYarnLock := fileExists(filepath.Join(teamsTokenDir, "yarn.lock")); hasYarnLock && commandExists("yarn") {
			if !fileExists(filepath.Join(teamsTokenDir, "node_modules")) {
				installCmd := exec.Command("yarn", "install")
				installCmd.Dir = teamsTokenDir
				installOutput, installErr := installCmd.CombinedOutput()
				if installErr != nil {
					return nil, "", fmt.Errorf("teams-token yarn install failed: %v (output: %s)", installErr, strings.TrimSpace(string(installOutput)))
				}
			}
			return buildElectronNodeCommand(teamsTokenDir), electronNodeCommandString(teamsTokenDir), nil
		}

		if commandExists("npm") {
			if !fileExists(filepath.Join(teamsTokenDir, "node_modules")) {
				installCmd := exec.Command("npm", "install", "--no-audit", "--no-fund")
				installCmd.Dir = teamsTokenDir
				installOutput, installErr := installCmd.CombinedOutput()
				if installErr != nil {
					return nil, "", fmt.Errorf("teams-token npm install failed: %v (output: %s)", installErr, strings.TrimSpace(string(installOutput)))
				}
			}
			return buildElectronNodeCommand(teamsTokenDir), electronNodeCommandString(teamsTokenDir), nil
		}
		return nil, "", fmt.Errorf("%s is a Node project, but neither yarn nor npm is installed", teamsTokenDir)
	}

	return nil, "", fmt.Errorf("%s exists but has no supported runner (binary/go.mod/package.json)", teamsTokenDir)
}

func buildElectronNodeCommand(teamsTokenDir string) *exec.Cmd {
	cmd := exec.Command("bash", "-lc", electronNodeCommandString(teamsTokenDir))
	cmd.Dir = teamsTokenDir
	return cmd
}

func electronNodeCommandString(teamsTokenDir string) string {
	electronBin := filepath.Join(teamsTokenDir, "node_modules", ".bin", "electron")
	electronFlags := []string{
		"--no-sandbox",
		"--disable-setuid-sandbox",
		"--disable-seccomp-filter-sandbox",
		"--disable-gpu",
		"--disable-dev-shm-usage",
	}
	electronEnv := "ELECTRON_DISABLE_SANDBOX=1"
	if useWaylandNativeElectron() {
		electronFlags = append(electronFlags, "--enable-features=UseOzonePlatform", "--ozone-platform=wayland")
		electronEnv = "ELECTRON_DISABLE_SANDBOX=1 ELECTRON_OZONE_PLATFORM_HINT=wayland"
	}
	return strings.Join([]string{
		"yarn run build",
		fmt.Sprintf("%s %s %s ./dist/main.js", electronEnv, shQuote(electronBin), strings.Join(electronFlags, " ")),
	}, " && ")
}

func wrapWithTermEverything(cmdStr string, teamsTokenDir string) (*exec.Cmd, bool, string, error) {
	decision := decideTermEverything()
	if decision.err != nil {
		return nil, false, decision.note, decision.err
	}
	if !decision.use {
		return nil, false, decision.note, nil
	}

	runCmd := fmt.Sprintf("cd %s && %s", shQuote(teamsTokenDir), cmdStr)
	wrappedCmd := runCmd
	if !useWaylandNativeElectron() {
		var wrapErr error
		wrappedCmd, wrapErr = wrapWithXwayland(runCmd)
		if wrapErr != nil {
			return nil, false, decision.note, wrapErr
		}
	} else {
		wrappedCmd = wrapWithLogging(runCmd)
	}

	if decision.binPath != "" {
		return exec.Command(decision.binPath, "--support-old-apps", "--", wrappedCmd), true, decision.note, nil
	}

	cmd := exec.Command("go", "run", ".", "--support-old-apps", "--", wrappedCmd)
	cmd.Dir = decision.termEverythingDir
	return cmd, true, decision.note, nil
}

func useWaylandNativeElectron() bool {
	val := strings.ToLower(strings.TrimSpace(os.Getenv("TEAMS_CLI_ELECTRON_OZONE")))
	return val == "1" || val == "true" || val == "yes"
}

type termEverythingDecision struct {
	note              string
	use               bool
	err               error
	binPath           string
	termEverythingDir string
}

func decideTermEverything() termEverythingDecision {
	termEverythingDir := findTermEverythingDir()
	if disableTermEverything() {
		return termEverythingDecision{note: "term.everything disabled", use: false, termEverythingDir: termEverythingDir}
	}

	if termEverythingDir == "" {
		return termEverythingDecision{note: "term.everything not found", use: false, termEverythingDir: "term.everything"}
	}

	if binPath := findTermEverythingBinary(termEverythingDir); binPath != "" {
		return termEverythingDecision{
			note:              "term.everything enabled (binary)",
			use:               true,
			binPath:           binPath,
			termEverythingDir: termEverythingDir,
		}
	}

	if commandExists("go") {
		return termEverythingDecision{
			note:              "term.everything enabled (go run)",
			use:               true,
			termEverythingDir: termEverythingDir,
		}
	}

	return termEverythingDecision{
		note:              "term.everything unavailable (no binary and Go missing)",
		use:               false,
		err:               fmt.Errorf("term.everything is present but no runnable binary found and Go is not installed"),
		termEverythingDir: termEverythingDir,
	}
}

func termEverythingNote() string {
	return decideTermEverything().note
}

func candidateSearchRoots() []string {
	roots := []string{}
	if wd, err := os.Getwd(); err == nil && wd != "" {
		roots = append(roots, wd)
	}
	if exeDir := executableDir(); exeDir != "" {
		roots = append(roots, exeDir)
	}
	if repoRoot := findRepoRoot(roots); repoRoot != "" {
		roots = append(roots, repoRoot)
	}
	return uniqStrings(roots)
}

func findTermEverythingDir() string {
	searchRoots := candidateSearchRoots()
	for _, root := range searchRoots {
		path := filepath.Join(root, "term.everything")
		info, err := os.Stat(path)
		if err == nil && info.IsDir() {
			return path
		}
	}
	return ""
}

func executableDir() string {
	exe, err := os.Executable()
	if err != nil || exe == "" {
		return ""
	}
	if resolved, err := filepath.EvalSymlinks(exe); err == nil && resolved != "" {
		exe = resolved
	}
	return filepath.Dir(exe)
}

func findRepoRoot(candidates []string) string {
	for _, start := range candidates {
		for dir := start; dir != "" && dir != "/"; dir = filepath.Dir(dir) {
			if fileExists(filepath.Join(dir, ".gitmodules")) {
				return dir
			}
		}
	}
	return ""
}

func uniqStrings(values []string) []string {
	seen := make(map[string]struct{}, len(values))
	out := make([]string, 0, len(values))
	for _, v := range values {
		if v == "" {
			continue
		}
		if _, ok := seen[v]; ok {
			continue
		}
		seen[v] = struct{}{}
		out = append(out, v)
	}
	return out
}

func disableTermEverything() bool {
	val := strings.ToLower(strings.TrimSpace(os.Getenv("TEAMS_CLI_DISABLE_TERM_EVERYTHING")))
	return val == "1" || val == "true" || val == "yes"
}

func findTermEverythingBinary(termEverythingDir string) string {
	pattern := filepath.Join(termEverythingDir, "dist", "*", "term.everything*")
	matches, _ := filepath.Glob(pattern)
	for _, match := range matches {
		if info, err := os.Stat(match); err == nil && !info.IsDir() && info.Mode()&0o111 != 0 {
			return match
		}
	}
	return ""
}

func wrapWithXwayland(runCmd string) (string, error) {
	if !commandExists("Xwayland") {
		return "", fmt.Errorf("Xwayland not found (install xorg-x11-server-Xwayland)")
	}
	if !commandExists("matchbox-window-manager") {
		return "", fmt.Errorf("matchbox-window-manager not found (install matchbox-window-manager)")
	}

	displayNum, err := findFreeXDisplay()
	if err != nil {
		return "", err
	}
	display := fmt.Sprintf(":%d", displayNum)
	socketPath := fmt.Sprintf("/tmp/.X11-unix/X%d", displayNum)
	logPath := fmt.Sprintf("/tmp/teams-cli-xwayland-%d.log", displayNum)

	// Start Xwayland + WM, wait for socket, run command, then cleanup.
	script := strings.Join([]string{
		fmt.Sprintf("exec >%s 2>&1", shQuote(logPath)),
		"set -x",
		fmt.Sprintf("Xwayland %s -retro & xw_pid=$!", display),
		"trap 'kill $xw_pid $wm_pid' EXIT",
		fmt.Sprintf("for i in $(seq 1 50); do [ -S %s ] && break; sleep 0.1; done", shQuote(socketPath)),
		fmt.Sprintf("[ -S %s ] || { echo \"Xwayland did not create socket %s\"; exit 1; }", shQuote(socketPath), socketPath),
		fmt.Sprintf("export DISPLAY=%s", display),
		fmt.Sprintf("matchbox-window-manager -display %s & wm_pid=$!", display),
		"export ELECTRON_DISABLE_SANDBOX=1",
		runCmd,
	}, "; ")

	return script, nil
}

func wrapWithLogging(runCmd string) string {
	// Capture output from the wrapped command for debugging.
	return strings.Join([]string{
		"log=/tmp/teams-cli-auth-$$.log",
		"echo \"auth refresh log: $log\"",
		"exec >$log 2>&1",
		"set -x",
		runCmd,
	}, "; ")
}

func findFreeXDisplay() (int, error) {
	for i := 5; i < 100; i++ {
		path := fmt.Sprintf("/tmp/.X11-unix/X%d", i)
		if _, err := os.Stat(path); err != nil {
			if os.IsNotExist(err) {
				return i, nil
			}
			return 0, err
		}
	}
	return 0, fmt.Errorf("no free X display found in /tmp/.X11-unix")
}

func shQuote(value string) string {
	if value == "" {
		return "''"
	}
	return "'" + strings.ReplaceAll(value, "'", "'\"'\"'") + "'"
}

func fileExists(path string) bool {
	info, err := os.Stat(path)
	return err == nil && info != nil
}

func commandExists(name string) bool {
	_, err := exec.LookPath(name)
	return err == nil
}

func normalizeComposeColorName(v string) string {
	switch strings.ToLower(strings.TrimSpace(v)) {
	case "midnight", "navy", "dark_blue", "slate":
		return strings.ToLower(strings.TrimSpace(v))
	default:
		return "midnight"
	}
}

func normalizeAuthorColorName(v string) string {
	switch strings.ToLower(strings.TrimSpace(v)) {
	case "blue", "yellow", "green", "cyan", "white":
		return strings.ToLower(strings.TrimSpace(v))
	default:
		return "blue"
	}
}
