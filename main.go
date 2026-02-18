package main

import (
	"os"

	"github.com/rivo/tview"
	"github.com/sirupsen/logrus"
)

func main() {
	app := tview.NewApplication()
	logger := logrus.New()
	logger.SetFormatter(&logrus.TextFormatter{
		FullTimestamp: true,
	})
	logger.SetLevel(logrus.DebugLevel)

	logFile := "teams-cli.log"
	f, err := os.OpenFile(logFile, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0o644)
	if err == nil {
		logger.SetOutput(f)
	}
	logger.WithFields(logrus.Fields{
		"log_file": logFile,
		"pid":      os.Getpid(),
	}).Info("teams-cli starting")

	state := AppState{
		app:    app,
		logger: logger,
	}

	state.createApp()
	if err = app.EnableMouse(true).Run(); err != nil {
		logger.WithError(err).Fatal("application exited with error")
	}
}
