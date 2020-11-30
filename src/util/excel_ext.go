package util

import (
	"fmt"
	"math"
	"strconv"
)

type TableListSheetExcelUtil struct {
	excelUtil             *ExcelUtil
	tableListSheetName    string //表清单页的sheet名
	tableNameColCharIndex string //表名列的列号 如"A"
	tableNameColIndex     int    //表名列的列数 如0
	hyperLinkStyleId      int    //超链接样式id
}

//ColNames 表清单sheet的列名
//tableColName 表名列的列名
func NewTableListSheetExcelUtil(filePath string, tableListSheetName string, ColNames []string, tableColName string) *TableListSheetExcelUtil {
	excelUtil := NewExcelUtil(filePath)
	tableNameColIndex := func() int {
		for i := 0; i < len(ColNames); i++ {
			if ColNames[i] == tableColName {
				return i
			}
		}
		panic(fmt.Sprintf("%s does not exist in %s", tableColName, ColNames))
	}()
	styleId, err := excelUtil.excel.NewStyle(`{"font":{"underline":"single","color":"#0066CC"}}`)
	PanicIfErr(err)
	return &TableListSheetExcelUtil{
		excelUtil:             excelUtil,
		tableListSheetName:    tableListSheetName,
		tableNameColIndex:     tableNameColIndex,
		tableNameColCharIndex: string(rune('A' + tableNameColIndex)),
		hyperLinkStyleId:      styleId,
	}
}
func (teu *TableListSheetExcelUtil) GetExcelUtil() *ExcelUtil {
	return teu.excelUtil
}

//列表页设置跳转到表的超链接
//columnName是表名所在的列名，即sql文件里设置的别名或字段名
func (teu *TableListSheetExcelUtil) SetHyperLinksToTableSheet() {
	excel := teu.excelUtil.excel
	rows, err := excel.Rows(teu.tableListSheetName)
	PanicIfErr(err)

	//第一行是列名，不设置超链接
	if !rows.Next() {
		return
	}
	rowNum := 2
	for rows.Next() {
		tableName := rows.Columns()[teu.tableNameColIndex]
		axis := teu.tableNameColCharIndex + strconv.Itoa(rowNum)

		hyperLinkSheetName := tableName
		if len(tableName) > 31 {
			fmt.Printf("[WARNING] the size of table name \"%s\" is large than 31,will cut off by excel",
				tableName)
			hyperLinkSheetName = tableName[0:31]
		}
		link := hyperLinkSheetName + "!A1"

		excel.SetCellHyperLink(teu.tableListSheetName, axis, link, "Location")
		excel.SetCellStyle(teu.tableListSheetName, axis, axis, teu.hyperLinkStyleId)
		rowNum++
	}
}

//各个表设置跳转到表清单sheet的超链接
func (teu *TableListSheetExcelUtil) SetHyperLinksToTableList() {
	excel := teu.excelUtil.excel
	for _, sheetName := range excel.GetSheetMap() {
		if sheetName == teu.tableListSheetName {
			continue
		}
		hyperLinkSheetName := sheetName[0:int(math.Min(float64(len(sheetName)), 31))]
		excel.SetCellHyperLink(hyperLinkSheetName, "A1", teu.tableListSheetName+"!A1", "Location")
		cellValue := excel.GetCellValue(sheetName, "A1")
		excel.SetCellValue(sheetName, "A1", cellValue+"(回到首页)")
		excel.SetCellStyle(sheetName, "A1", "A1", teu.hyperLinkStyleId)
	}
}
