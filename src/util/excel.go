package util

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"math"
	"strconv"
	"sync"
)

type ExcelUtil struct {
	excel    *excelize.File
	filePath string
	mu       sync.Mutex

	tableListSheetName    string //表清单页的sheet名
	tableNameColCharIndex string //表名列的列号 如"A"
	tableNameColIndex     int    //表名列的列数 如0
	hyperLinkStyleId      int    //超链接样式id
}

func (eu *ExcelUtil) NewSheet(sheetName string, columnNames []string, rows [][]string) {
	eu.mu.Lock()
	defer eu.mu.Unlock()
	//如果index是2，则是除了默认的Sheet1之外新建的第一个Sheet
	if eu.excel.NewSheet(sheetName) == 2 {
		eu.excel.DeleteSheet("Sheet1")
	}
	maxColStrLen := map[rune]int{}

	for i, name := range columnNames {
		//列宽首先考虑字段名
		eu.excel.SetCellValue(sheetName, string(rune('A'+i))+"1", name)
		maxColStrLen[rune('A'+i)] = len(name)
	}
	for i, row := range rows {
		for j, s := range row {
			eu.excel.SetCellValue(sheetName, string(rune('A'+j))+strconv.Itoa(i+2), s)
			//统计这一列中字符串长度的最大值
			if v, ok := maxColStrLen[rune('A'+j)]; !ok || v < len(s) {
				maxColStrLen[rune('A'+j)] = len(s)
			}
		}
	}

	//设置列宽防止显示不完整
	for col, strLen := range maxColStrLen {
		eu.excel.SetColWidth(sheetName, string(col), string(col),
			math.Max(float64(strLen)*1.3, 10.0))
	}
}
func (eu *ExcelUtil) Save() {
	eu.mu.Lock()
	defer eu.mu.Unlock()
	eu.excel.SetActiveSheet(2)
	PanicIfErr(eu.excel.SaveAs(eu.filePath))
}

//ColNames 表清单sheet的列名
//tableColName 表名列的列名
func NewExcelUtil(filePath string, tableListSheetName string, ColNames []string, tableColName string) *ExcelUtil {
	excel := excelize.NewFile()
	tableNameColIndex := func() int {
		for i := 0; i < len(ColNames); i++ {
			if ColNames[i] == tableColName {
				return i
			}
		}
		panic(fmt.Sprintf("%s does not exist in %s", tableColName, ColNames))
	}()
	//预设超链接样式
	styleId, err := excel.NewStyle(`{"font":{"underline":"single","color":"#0066CC"}}`)
	PanicIfErr(err)
	return &ExcelUtil{
		excel:                 excel,
		filePath:              filePath,
		mu:                    sync.Mutex{},
		tableListSheetName:    tableListSheetName,
		tableNameColIndex:     tableNameColIndex,
		tableNameColCharIndex: string(rune('A' + tableNameColIndex)),
		hyperLinkStyleId:      styleId,
	}
}

//列表页设置跳转到表的超链接
//columnName是表名所在的列名，即sql文件里设置的别名或字段名
func (eu *ExcelUtil) SetHyperLinksToTableSheet() {
	excel := eu.excel
	rows, err := excel.Rows(eu.tableListSheetName)
	PanicIfErr(err)

	//第一行是列名，不设置超链接
	if !rows.Next() {
		return
	}
	rowNum := 2
	for rows.Next() {
		tableName := rows.Columns()[eu.tableNameColIndex]
		axis := eu.tableNameColCharIndex + strconv.Itoa(rowNum)

		hyperLinkSheetName := tableName
		if len(tableName) > 31 {
			fmt.Printf("[WARNING] the size of table name \"%s\" is large than 31,will cut off by excel",
				tableName)
			hyperLinkSheetName = tableName[0:31]
		}
		link := hyperLinkSheetName + "!A1"

		excel.SetCellHyperLink(eu.tableListSheetName, axis, link, "Location")
		excel.SetCellStyle(eu.tableListSheetName, axis, axis, eu.hyperLinkStyleId)
		rowNum++
	}
}

//各个表设置跳转到表清单sheet的超链接
func (eu *ExcelUtil) SetHyperLinksToTableList() {
	excel := eu.excel
	for _, sheetName := range excel.GetSheetMap() {
		if sheetName == eu.tableListSheetName {
			continue
		}
		hyperLinkSheetName := sheetName[0:int(math.Min(float64(len(sheetName)), 31))]
		excel.SetCellHyperLink(hyperLinkSheetName, "A1", eu.tableListSheetName+"!A1", "Location")
		cellValue := excel.GetCellValue(sheetName, "A1")
		excel.SetCellValue(sheetName, "A1", cellValue+"(回到首页)")
		excel.SetCellStyle(sheetName, "A1", "A1", eu.hyperLinkStyleId)
	}
}
