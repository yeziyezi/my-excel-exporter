package util

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

type ExcelUtil struct {
	excel    *excelize.File
	filePath string
}

func NewExcelUtil(filePath string) *ExcelUtil {
	excel := excelize.NewFile()

	return &ExcelUtil{
		excel:    excel,
		filePath: filePath,
	}
}
func (eu *ExcelUtil) NewSheet(sheetName string, columnNames []string, rows [][]string) {
	sheetIndex := eu.excel.NewSheet(sheetName)
	//如果index是2，则是除了默认的Sheet1之外新建的第一个Sheet
	//把Sheet1删掉同时将这个Sheet置为默认显示Sheet
	if sheetIndex == 2 {
		eu.excel.DeleteSheet("Sheet1")
		eu.excel.SetActiveSheet(sheetIndex)
	}
	for i, name := range columnNames {
		eu.excel.SetCellValue(sheetName, string(rune('A'+i))+"1", name)
	}
	//eu.excel.SetColWidth(sheetName, string('A'), string(rune('A'+len(columnNames)-1)), 50)
	for i, row := range rows {
		for j, s := range row {
			eu.excel.SetCellValue(sheetName, string(rune('A'+j))+strconv.Itoa(i+2), s)
		}
	}
}
func (eu *ExcelUtil) Save() {
	PanicIfErr(eu.excel.SaveAs(eu.filePath))
}
