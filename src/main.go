package main

import (
	"db-excel-export-helper/src/util"
	_ "github.com/go-sql-driver/mysql"
	"os"
)

type newSheetParam struct {
	tableName string
	rows      [][]string
}

func main() {
	cwd, _ := os.Getwd()
	config := util.ReadConfig(cwd + "/conf/config.json")

	db := util.GetDB(config)
	defer func() { _ = db.Close() }()

	listQuery := util.NewQuery(cwd+"/conf/tables.sql", db)
	tableQuery := util.NewQuery(cwd+"/conf/table-struct.sql", db)

	listRows := listQuery.QueryAll(config.Schema)

	e := util.NewExcelUtil(config, listQuery.GetColumnNames(), tableQuery.GetColumnNames())
	e.NewSheet(config.ListSheetName, listRows)

	var tableNames []string
	for _, row := range listRows {
		tableNames = append(tableNames, row[0])
	}
	c := make(chan newSheetParam)
	//限制查询并发数避免造成数据库压力
	queryLimitChan := make(chan int, config.ConcurrentNum)
	for _, tableName := range tableNames {
		tableName := tableName
		go func() {
			queryLimitChan <- 1
			defer func() { <-queryLimitChan }()
			rows := tableQuery.QueryAll(config.Schema, tableName)
			c <- newSheetParam{tableName: tableName, rows: rows}
			util.Log.Infoln(tableName + "...ok")
		}()
	}
	for range tableNames {
		nsp := <-c
		e.NewSheet(nsp.tableName, nsp.rows)
	}

	util.Log.Infof("%d tables done\n", len(tableNames))
	util.Log.Infof("writing into %s...", config.ExcelPath)
	e.End()
	util.Log.Infoln("success")
	util.WaitForEnterAndExit()
}
