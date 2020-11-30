package main

import (
	"db-excel-export-helper/src/util"
	"fmt"
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
	e.NewSheet(config.ListTabName, listRows)

	var tableNames []string
	for _, row := range listRows {
		tableNames = append(tableNames, row[0])
	}
	c := make(chan newSheetParam)
	for _, tableName := range tableNames {
		tableName := tableName
		go func() {
			rows := tableQuery.QueryAll(config.Schema, tableName)
			c <- newSheetParam{tableName: tableName, rows: rows}
			fmt.Println(tableName + "...ok")
		}()
	}
	for range tableNames {
		nsp := <-c
		e.NewSheet(nsp.tableName, nsp.rows)
	}

	fmt.Printf("%d tables done\n", len(tableNames))
	fmt.Print("setting hyperlinks...")
	e.SetHyperLinksInListSheet()
	e.SetHyperLinksInTableSheets()
	fmt.Println("ok")
	fmt.Printf("writing into %s...", config.ExcelPath)
	e.Save()
	fmt.Println("success")
}
