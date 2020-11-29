package util

import (
	"database/sql"
	"encoding/json"
	"fmt"
	_ "github.com/go-sql-driver/mysql"
	"io/ioutil"
	"time"
)

//https://github.com/go-sql-driver/mysql/wiki/Examples
//https://www.xiexiaodong7.com/archives/17
type Config struct {
	Driver       string `json:"driver"`
	Username     string `json:"username"`
	Password     string `json:"password"`
	Host         string `json:"host"`
	Port         string `json:"port"`
	Schema       string `json:"schema"`
	ListTabName  string `json:"listTabName"` //第一个tab用于存放表名清单
	ExcelPath    string `json:"excelPath"`
	TableColName string `json:"tableColName"` //表名列的列名
}

func ReadConfig(path string) *Config {
	nByte, err := ioutil.ReadFile(path)
	PanicIfErr(err)
	var config Config
	err = json.Unmarshal(nByte, &config)
	PanicIfErr(err)
	return &config
}
func PanicIfErr(err error) {
	if err != nil {
		panic(err)
	}
}
func GetDB(config *Config) *sql.DB {
	confString := fmt.Sprintf("%s:%s@(%s:%s)/%s",
		config.Username, config.Password, config.Host, config.Port, config.Schema)
	db, err := sql.Open(config.Driver, confString)
	if err != nil {
		panic(err)
	}
	db.SetConnMaxLifetime(time.Minute * 3)
	db.SetMaxOpenConns(10)
	db.SetMaxIdleConns(10)
	err = db.Ping()
	if err != nil {
		panic(err)
	}
	return db
}
