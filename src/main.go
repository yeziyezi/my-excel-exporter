package main

import (
	"db-excel-export-helper/src/util"
	"fmt"
	_ "github.com/go-sql-driver/mysql"
	"os"
)

type newSheetParam struct {
	tableName   string
	columnNames []string
	rows        [][]string
}

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
	c := make(chan newSheetParam)
	for _, tableName := range tableNames {
		tableName := tableName
		go func() {
			rows := query.QueryAll(config.Schema, tableName)
			c <- newSheetParam{tableName: tableName, columnNames: query.GetColumnNames(), rows: rows}
			fmt.Println(tableName + "...ok")
		}()
	}
	for range tableNames {
		nsp := <-c
		e.NewSheet(nsp.tableName, nsp.columnNames, nsp.rows)
	}

	fmt.Printf("%d tables done\n", len(tableNames))
	fmt.Printf("writing into %s...", config.ExcelPath)
	e.Save()
	fmt.Println("success")
}
