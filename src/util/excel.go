package util

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"math"
	"strconv"
	"sync"
)

//ExcelUtil 封装好的Excel工具类
type ExcelUtil struct {
	excel                 *excelize.File
	mu                    sync.Mutex
	filePath              string            //保存excel文件的路径
	tableListSheetName    string            //表清单页的sheet名
	tableNameColCharIndex string            //表清单sheet中表名列的excel列号 如"A"
	tableNameColIndex     int               //表清单sheet中表名列的数字列号 如0
	hyperLinkStyleId      int               //超链接样式id
	hyperLinkTitleStyleId int               //超链接标题样式id
	tableSheetCols        []string          //表结构sheet的列名数组
	listSheetCols         []string          //表清单sheet列名数组
	sheetTableMap         map[string]string //sheetName与tableName的映射
	tableSheetMap         map[string]string //tableName与sheetName的映射
}

func (eu *ExcelUtil) NewSheet(name string, rows [][]string) {
	eu.mu.Lock()
	defer eu.mu.Unlock()

	var tableName string
	var sheetName string

	isListSheet := name == eu.tableListSheetName
	//判断是否为表清单sheet
	if isListSheet {
		sheetName = name
	} else {
		tableName = name
		sheetName = "Table" + strconv.Itoa(len(eu.excel.GetSheetMap()))
		eu.sheetTableMap[sheetName] = tableName
		eu.tableSheetMap[tableName] = sheetName
	}

	//如果index是2，则是除了默认的Sheet1之外新建的第一个Sheet
	if eu.excel.NewSheet(sheetName) == 2 {
		eu.excel.DeleteSheet("Sheet1")
	}

	//设置表结构sheet A1为表名
	if !isListSheet {
		eu.excel.SetCellValue(sheetName, "A1", tableName)
	}

	maxColStrLen := map[rune]int{}

	//列名行
	var colTitleRow int
	if isListSheet {
		colTitleRow = 1
	} else {
		colTitleRow = 2
	}

	var cols []string
	if isListSheet {
		cols = eu.listSheetCols
	} else {
		cols = eu.tableSheetCols
	}
	for i, name := range cols {
		eu.excel.SetCellValue(sheetName, string(rune('A'+i))+strconv.Itoa(colTitleRow), name)
		//列宽首先考虑字段名
		maxColStrLen[rune('A'+i)] = len(name)
	}

	for i, row := range rows {
		for j, s := range row {
			eu.excel.SetCellValue(sheetName, string(rune('A'+j))+strconv.Itoa(i+colTitleRow+1), s)
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

//@param config *Config 配置项
//@param listSheetCols []string 表清单sheet列名
//@param tableSheetCols []string 表结构sheet列名
func NewExcelUtil(config *Config, listSheetCols []string, tableSheetCols []string) *ExcelUtil {
	excel := excelize.NewFile()
	tableNameColIndex := func() int {
		for i := 0; i < len(listSheetCols); i++ {
			if listSheetCols[i] == config.TableColName {
				return i
			}
		}
		panic(fmt.Sprintf("%s does not exist in %s", config.TableColName, listSheetCols))
	}()
	//预设超链接样式
	hyperLinkStyleId, err := excel.NewStyle(`{"font":{"underline":"single","color":"#0066CC"}}`)
	hyperLinkTitleStyleId, err := excel.NewStyle(`
{
    "font": {
        "underline": "single",
        "color": "#0066CC"
    },
    "alignment": {
        "horizontal": "center"
    }
}
`)

	PanicIfErr(err)
	return &ExcelUtil{
		excel:                 excel,
		mu:                    sync.Mutex{},
		filePath:              config.ExcelPath,
		tableListSheetName:    config.ListSheetName,
		tableNameColIndex:     tableNameColIndex,
		tableNameColCharIndex: string(rune('A' + tableNameColIndex)),
		hyperLinkStyleId:      hyperLinkStyleId,
		hyperLinkTitleStyleId: hyperLinkTitleStyleId,
		tableSheetCols:        tableSheetCols,
		sheetTableMap:         map[string]string{},
		tableSheetMap:         map[string]string{},
		listSheetCols:         listSheetCols,
	}
}

//列表页设置跳转到表的超链接
func (eu *ExcelUtil) SetHyperLinksInListSheet() {
	eu.mu.Lock()
	defer eu.mu.Unlock()

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
		link := eu.tableSheetMap[tableName] + "!A1"

		excel.SetCellHyperLink(eu.tableListSheetName, axis, link, "Location")
		excel.SetCellStyle(eu.tableListSheetName, axis, axis, eu.hyperLinkStyleId)
		rowNum++
	}
}

//各个表设置跳转到表清单sheet的超链接
func (eu *ExcelUtil) SetHyperLinksInTableSheets() {
	eu.mu.Lock()
	defer eu.mu.Unlock()
	excel := eu.excel
	for _, sheetName := range excel.GetSheetMap() {
		if sheetName == eu.tableListSheetName {
			continue
		}
		hyperLinkSheetName := sheetName[0:int(math.Min(float64(len(sheetName)), 31))]
		excel.SetCellHyperLink(hyperLinkSheetName, "A1", eu.tableListSheetName+"!A1", "Location")
		cellValue := excel.GetCellValue(sheetName, "A1")
		excel.SetCellValue(sheetName, "A1", cellValue+"(回到首页)")
		excel.SetCellStyle(sheetName, "A1", "A1", eu.hyperLinkTitleStyleId)
		excel.MergeCell(sheetName, "A1", string(rune('A'+len(eu.tableSheetCols)-1))+"1")
	}
}
