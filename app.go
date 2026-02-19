package main

import (
	"bytes"
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

	authRefreshMu sync.Mutex
}

type conversationRef struct {
	ids        []string
	title      string
	chatKey    string
	isFavorite bool
}

func (s *AppState) createApp() {
	s.logger.Debug("creating application pages and components")
	s.pages = tview.NewPages()
	s.components = map[string]tview.Primitive{}
	s.chatFavorites = map[string]bool{}

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
		SetPlaceholder("Press i to compose, f to toggle favorite chat, Enter to send, Esc to return")
	composeView.SetFieldWidth(0)
	composeView.SetBorder(true)
	composeView.SetTitle("Compose")
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
	s.pages.SwitchToPage(PageError)
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

	s.components[TvError] = p

	return tview.NewFlex().
		AddItem(nil, 0, 1, false).
		AddItem(tview.NewFlex().SetDirection(tview.FlexRow).
			AddItem(nil, 0, 1, false).
			AddItem(p, 10, 1, false).
			AddItem(nil, 0, 1, false), 60, 1, false).
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
		chatKey := chatFavoriteKey(chat.Id, candidateIDs)
		isFavorite := s.chatIsFavorite(chat, chatKey)
		chatNode.SetReference(conversationRef{
			ids:        candidateIDs,
			title:      chatName,
			chatKey:    chatKey,
			isFavorite: isFavorite,
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
	s.logger.WithFields(logrus.Fields{
		"chat_nodes_count":      len(favoritesNode.GetChildren()) + len(recentNode.GetChildren()),
		"favorites_nodes_count": len(favoritesNode.GetChildren()),
		"recent_nodes_count":    len(recentNode.GetChildren()),
	}).Debug("chat tree nodes prepared")

	treeView.SetSelectedFunc(func(node *tview.TreeNode) {
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
			s.logger.WithFields(logrus.Fields{
				"target":        "chat",
				"display_name":  ref.title,
				"candidate_ids": strings.Join(ref.ids, ","),
			}).Info("loading conversation")
			s.components[ViChat].(*tview.List).
				SetTitle(ref.title).
				SetBorder(true).
				SetTitleAlign(tview.AlignCenter)
			s.setActiveConversation(node, ref.ids, ref.title)
			go s.loadConversationsByIDs(node, ref.ids, ref.title)
		}
	})
	treeView.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		if event.Key() == tcell.KeyRune && (event.Rune() == 'f' || event.Rune() == 'F') {
			if s.toggleFavoriteForCurrentNode(treeView, chatsNode, favoritesNode, recentNode) {
				return nil
			}
			return event
		}
		if event.Key() == tcell.KeyRune && (event.Rune() == 'u' || event.Rune() == 'U') {
			go s.refreshAllChatLabels(chatsNode)
			return nil
		}
		if event.Key() == tcell.KeyRune && (event.Rune() == 'i' || event.Rune() == 'I') {
			ids, _, _ := s.getActiveConversation()
			if len(ids) == 0 {
				return event
			}
			s.app.SetFocus(composeView)
			return nil
		}
		return event
	})
	composeView.SetDoneFunc(func(key tcell.Key) {
		if key == tcell.KeyEscape {
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
		composeView.SetText("")
		go s.sendMessageAndRefresh(ids, title, messageText, selectedNode)
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
	s.logger.Info("main window ready")
}

func (s *AppState) setActiveConversation(selectedNode *tview.TreeNode, conversationIDs []string, title string) {
	ids := normalizeConversationIDs(conversationIDs)
	s.activeConversationMu.Lock()
	s.activeConversationIDs = ids
	s.activeConversationTitle = title
	s.activeConversationNode = selectedNode
	s.activeConversationMu.Unlock()
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
	return next
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
		s.app.QueueUpdateDraw(func() {
			node.SetText(title)
			node.SetReference(ref)
		})
	}

	s.logger.WithFields(logrus.Fields{
		"attempted": attempted,
		"updated":   updated,
	}).Info("finished refreshing chat titles")
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
			output += fmt.Sprintf("\t%v\n", text)
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

func (s *AppState) sendMessageAndRefresh(conversationIDs []string, displayName, content string, selectedNode *tview.TreeNode) {
	err := s.sendMessage(conversationIDs, content)
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

func formatOutgoingHTML(content string) string {
	lines := strings.Split(content, "\n")
	for i := range lines {
		lines[i] = strings.TrimSpace(lines[i])
	}
	body := strings.Join(lines, "<br/>")
	return "<div><div>" + body + "</div></div>"
}

func (s *AppState) sendMessage(conversationIDs []string, content string) error {
	ids := normalizeConversationIDs(conversationIDs)
	if len(ids) == 0 {
		return fmt.Errorf("no conversation id available")
	}

	payload := map[string]interface{}{
		"content":         formatOutgoingHTML(content),
		"messagetype":     "RichText/Html",
		"contenttype":     "text",
		"clientmessageid": strconv.FormatInt(time.Now().UnixNano(), 10),
		"amsreferences":   []string{},
		"properties":      map[string]interface{}{},
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
		selectedNode.SetText(displayName)
		if ref, ok := selectedNode.GetReference().(conversationRef); ok {
			ref.title = displayName
			selectedNode.SetReference(ref)
		}
	}
	s.components[ViChat].(*tview.List).
		SetTitle(displayName).
		SetBorder(true).
		SetTitleAlign(tview.AlignCenter)

	// Clear chat
	chatList := s.components[ViChat].(*tview.List)
	chatList.Clear()
	s.logger.WithFields(logrus.Fields{
		"display_name":   displayName,
		"messages_count": len(messages),
	}).Debug("rendering messages")
	for _, message := range messages {
		author := strings.TrimSpace(message.ImDisplayName)
		if author == "" {
			author = inferMessageAuthor(message, s.me)
		}
		chatList.AddItem(textMessage(message.Content), author, 0, nil)
	}
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

func isUnauthorizedError(err error) bool {
	if err == nil {
		return false
	}
	lower := strings.ToLower(strings.TrimSpace(err.Error()))
	return strings.Contains(lower, "401") || strings.Contains(lower, "unauthorized")
}

func (s *AppState) refreshAuthFromTeamsToken() error {
	s.authRefreshMu.Lock()
	defer s.authRefreshMu.Unlock()

	teamsTokenDir := "teams-token"
	info, err := os.Stat(teamsTokenDir)
	if err != nil || !info.IsDir() {
		return fmt.Errorf("optional %s directory not found", teamsTokenDir)
	}

	var cmd *exec.Cmd
	binaryPath := filepath.Join(teamsTokenDir, "teams-token")
	if binaryInfo, statErr := os.Stat(binaryPath); statErr == nil && !binaryInfo.IsDir() && binaryInfo.Mode()&0o111 != 0 {
		cmd = exec.Command("./teams-token")
		cmd.Dir = teamsTokenDir
	} else {
		if _, statErr = os.Stat(filepath.Join(teamsTokenDir, "go.mod")); statErr != nil {
			return fmt.Errorf("%s exists but has no teams-token binary or go.mod", teamsTokenDir)
		}
		cmd = exec.Command("go", "run", ".")
		cmd.Dir = teamsTokenDir
	}

	output, runErr := cmd.CombinedOutput()
	if runErr != nil {
		return fmt.Errorf("teams-token refresh command failed: %v (output: %s)", runErr, strings.TrimSpace(string(output)))
	}
	s.logger.Info("teams-token refresh command succeeded")

	newClient, newClientErr := teams_api.New()
	if newClientErr != nil {
		return fmt.Errorf("unable to reinitialize Teams client after token refresh: %v", newClientErr)
	}
	s.teamsClient = newClient
	return nil
}
