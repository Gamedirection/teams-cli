package main

import (
	"fmt"
	teams_api "github.com/fossteams/teams-api"
	"github.com/fossteams/teams-api/pkg/csa"
	"github.com/fossteams/teams-api/pkg/models"
	"github.com/sirupsen/logrus"
	"sort"
)

type TeamsState struct {
	teamsClient *teams_api.TeamsClient
	logger      *logrus.Logger

	conversations  *csa.ConversationResponse
	me             *models.User
	pinnedChannels []csa.ChannelId
	channelById    map[string]Channel
	teamById       map[string]*csa.Team
}

type Channel struct {
	*csa.Channel
	parent *csa.Team
}

func (s *TeamsState) init(client *teams_api.TeamsClient) error {
	if client == nil {
		return fmt.Errorf("client is nil")
	}
	if s.logger != nil {
		s.logger.Debug("teams state initialization started")
	}

	var err error
	s.me, err = client.GetMe()
	if err != nil {
		return fmt.Errorf("unable to get your profile: %v", err)
	}
	if s.logger != nil {
		s.logger.WithFields(logrus.Fields{
			"user_display_name": s.me.DisplayName,
			"user_oid":          s.me.ObjectId,
			"user_upn":          s.me.UserPrincipalName,
		}).Info("loaded current user profile")
	}

	s.pinnedChannels, err = client.GetPinnedChannels()
	if err != nil {
		return fmt.Errorf("unable to get pinned channels: %v", err)
	}
	if s.logger != nil {
		s.logger.WithField("pinned_channels_count", len(s.pinnedChannels)).Debug("loaded pinned channels")
	}

	s.conversations, err = client.GetConversations()
	if err != nil {
		return fmt.Errorf("unable to get conversations: %v", err)
	}
	if s.logger != nil {
		s.logger.WithFields(logrus.Fields{
			"teams_count":         len(s.conversations.Teams),
			"chats_count":         len(s.conversations.Chats),
			"private_feeds_count": len(s.conversations.PrivateFeeds),
		}).Info("loaded conversations payload")
	}

	// Sort Teams by Name
	sort.Sort(csa.TeamsByName(s.conversations.Teams))

	// Create maps
	s.teamById = map[string]*csa.Team{}
	s.channelById = map[string]Channel{}

	for _, t := range s.conversations.Teams {
		s.teamById[t.Id] = &t
		for _, c := range t.Channels {
			s.channelById[c.Id] = Channel{
				Channel: &c,
				parent:  &t,
			}
		}
	}
	if s.logger != nil {
		s.logger.WithFields(logrus.Fields{
			"team_map_count":    len(s.teamById),
			"channel_map_count": len(s.channelById),
		}).Debug("conversation indexes built")
	}

	return nil
}
