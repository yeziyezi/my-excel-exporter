package util

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"math"
	"strconv"
	"sync"
)

type ExcelUtil struct {
	excel    *excelize.File
	filePath string
	mu       sync.Mutex
}

func NewExcelUtil(filePath string) *ExcelUtil {
	return &ExcelUtil{
		excel:    excelize.NewFile(),
		filePath: filePath,
		mu:       sync.Mutex{},
	}
}
func (eu *ExcelUtil) NewSheet(sheetName string, columnNames []string, rows [][]string) {
	eu.mu.Lock()
	defer eu.mu.Unlock()
	sheetIndex := eu.excel.NewSheet(sheetName)
	//如果index是2，则是除了默认的Sheet1之外新建的第一个Sheet
	//把Sheet1删掉同时将这个Sheet置为默认显示Sheet
	if sheetIndex == 2 {
		eu.excel.DeleteSheet("Sheet1")
		eu.excel.SetActiveSheet(sheetIndex)
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
			if v, ok := maxColStrLen[rune('A'+j)]; !ok || v < len(s) {
				maxColStrLen[rune('A'+j)] = len(s)
			}
		}
	}

	//设置列宽防止显示不完整
	for col, strLen := range maxColStrLen {
		eu.excel.SetColWidth(sheetName, string(col), string(col), math.Max(float64(strLen)*1.3, 10.0))
	}
}
func (eu *ExcelUtil) Save() {
	eu.mu.Lock()
	defer eu.mu.Unlock()
	PanicIfErr(eu.excel.SaveAs(eu.filePath))
}
