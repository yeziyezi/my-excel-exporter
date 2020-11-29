package main

import (
	"db-excel-export-helper/src/util"
	_ "github.com/go-sql-driver/mysql"
	"os"
)

func main() {
	cwd, _ := os.Getwd()
	config := util.ReadConfig(cwd + "/conf/config.json")

	db := util.GetDB(config)
	defer func() { _ = db.Close() }()

	tableListQuery := util.NewQuery(cwd+"/conf/tables.sql", db)
	tableListRows := tableListQuery.QueryAll(config.Schema)

	e := util.NewExcelUtil(config.ExcelPath)
	e.NewSheet(config.ListTabName, tableListQuery.GetColumnNames(), tableListRows)

	query := util.NewQuery(cwd+"/conf/table-struct.sql", db)

	var tableNames []string
	for _, row := range tableListRows {
		tableNames = append(tableNames, row[0])
	}
	for _, tableName := range tableNames {
		rows := query.QueryAll(config.Schema, tableName)
		e.NewSheet(tableName, query.GetColumnNames(), rows)
	}

	e.Save()

}
