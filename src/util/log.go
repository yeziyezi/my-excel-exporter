package util

import (
	"fmt"
	"log"
	"os"
)

var Log *Logger

type Logger struct {
	fatalLogger *log.Logger
	infoLogger  *log.Logger
}

func (log *Logger) Fatal(err error) {
	_ = log.fatalLogger.Output(3, fmt.Sprint(err))
}
func (log *Logger) Infoln(s string) {
	log.infoLogger.Println(s)
}
func (log *Logger) Infof(format string, v ...interface{}) {
	log.infoLogger.Printf(format, v)
}
func init() {
	Log = new(Logger)
	Log.fatalLogger = log.New(os.Stdout, "ERROR ", log.LstdFlags|log.Lshortfile)
	Log.infoLogger = log.New(os.Stdout, "INFO  ", log.LstdFlags|log.Lshortfile)
}
