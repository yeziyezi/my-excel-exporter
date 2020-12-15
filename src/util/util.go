package util

import (
	"database/sql"
	"encoding/json"
	"fmt"
	_ "github.com/go-sql-driver/mysql"
	"io/ioutil"
	"os"
	"time"
)

//https://github.com/go-sql-driver/mysql/wiki/Examples
//https://www.xiexiaodong7.com/archives/17
type Config struct {
	Driver        string                 `json:"driver"`
	Username      string                 `json:"username"`
	Password      string                 `json:"password"`
	Host          string                 `json:"host"`
	Port          string                 `json:"port"`
	Schema        string                 `json:"schema"`
	ListSheetName string                 `json:"listSheetName"` //第一个tab用于存放表名清单
	ExcelPath     string                 `json:"excelPath"`
	TableColName  string                 `json:"tableColName"` //表名列的列名
	Style         map[string]interface{} `json:"style"`        //样式
}

func ReadConfig(path string) *Config {
	nByte, err := ioutil.ReadFile(path)
	ExitIfErr(err)
	var config Config
	err = json.Unmarshal(nByte, &config)
	ExitIfErr(err)
	return &config
}
func ExitIfErr(err error) {
	if err != nil {
		Log.Fatal(err)
		WaitForEnterAndExit()
	}
}
func GetDB(config *Config) *sql.DB {
	confString := fmt.Sprintf("%s:%s@(%s:%s)/%s",
		config.Username, config.Password, config.Host, config.Port, config.Schema)
	db, err := sql.Open(config.Driver, confString)
	ExitIfErr(err)
	db.SetConnMaxLifetime(time.Minute * 3)
	db.SetMaxOpenConns(10)
	db.SetMaxIdleConns(10)
	err = db.Ping()
	ExitIfErr(err)
	return db
}

//等待键入回车
func WaitForEnterAndExit() {
	fmt.Println("===================")
	fmt.Println("press ENTER to exit.")
	_, _ = fmt.Scanf("\n")
	os.Exit(1)
}
