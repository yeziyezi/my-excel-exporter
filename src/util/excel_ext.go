package util

import (
	"fmt"
	"strconv"
)

type TableListSheetExcelUtil struct {
	e                     *ExcelUtil
	tableListSheetName    string //表清单页的sheet名
	tableNameColCharIndex string //表名列的列号 如"A"
	tableNameColIndex     int    //表名列的列数 如0
}

//ColNames 表清单sheet的列名
//tableColName 表名列的列名
func NewTableListSheetExcelUtil(filePath string, tableListSheetName string, ColNames []string, tableColName string) *TableListSheetExcelUtil {
	tableNameColIndex := func() int {
		for i := 0; i < len(ColNames); i++ {
			if ColNames[i] == tableColName {
				return i
			}
		}
		panic(fmt.Sprintf("%s does not exist in %s", tableColName, ColNames))
	}()
	return &TableListSheetExcelUtil{
		e:                     NewExcelUtil(filePath),
		tableListSheetName:    tableListSheetName,
		tableNameColIndex:     tableNameColIndex,
		tableNameColCharIndex: string(rune('A' + tableNameColIndex)),
	}
}
func (teu *TableListSheetExcelUtil) GetExcelUtil() *ExcelUtil {
	return teu.e
}

//列表页设置跳转到表的超链接
//columnName是表名所在的列名，即sql文件里设置的别名或字段名
func (teu *TableListSheetExcelUtil) SetHyperLinksToTableSheet() {
	excel := teu.e.excel
	rows, err := excel.Rows(teu.tableListSheetName)
	PanicIfErr(err)
	rowNum := 1
	for rows.Next() {
		tableName := rows.Columns()[teu.tableNameColIndex]
		excel.SetCellHyperLink(teu.tableListSheetName,
			teu.tableNameColCharIndex+strconv.Itoa(rowNum),
			tableName+"!A1",
			"Location")
		rowNum++
	}
}
