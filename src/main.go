package main

import (
	"db-excel-export-helper/src/util"
	_ "github.com/go-sql-driver/mysql"
	"os"
)

//todo 调整表格输出的样式
//根据字段长度调整行宽
//首个sheet增加跳转到其他sheet的链接

//todo 报错提示友好化
//todo 打印进度信息
//todo 提交到GitHub
//todo 打zip包作为release
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
