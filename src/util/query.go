package util

import (
	"database/sql"
	"io/ioutil"
	"strings"
)

type QueryUtil struct {
	stmt        *sql.Stmt
	buf         []*[]byte
	iBuf        []interface{}
	columnNames []string //列名的别名或者列名，取决于查询是否使用了AS
}

func readStringFromFile(sqlFilePath string) string {
	//读取SQL文件中的SQL语句
	sqlBytes, err := ioutil.ReadFile(sqlFilePath)
	PanicIfErr(err)
	return string(sqlBytes)
}
func buildSqlStatement(sqlString string, db *sql.DB) *sql.Stmt {
	//构建预编译语句
	stmt, err := db.Prepare(sqlString)
	PanicIfErr(err)
	return stmt
}
func createBuffer(bufLen int) (buf []*[]byte, iBuf []interface{}) {
	//根据SQL文件中解析出的字段数创建一个等长度的切片作为每一行数据的buffer
	for i := 0; i < bufLen; i++ {
		var bt []byte
		buf = append(buf, &bt)
	}
	//由于Rows.Scan只接收...Interface{}，进行转换
	for _, t := range buf {
		iBuf = append(iBuf, t)
	}
	return
}
func NewQuery(sqlFilePath string, db *sql.DB) *QueryUtil {
	sqlString := readStringFromFile(sqlFilePath)
	stmt := buildSqlStatement(sqlString, db)
	columnNames := readColumnNames(sqlString)
	buf, iBuf := createBuffer(len(columnNames))
	return &QueryUtil{
		stmt:        stmt,
		buf:         buf,
		iBuf:        iBuf,
		columnNames: columnNames,
	}
}

func (qu *QueryUtil) QueryAll(param ...interface{}) (result [][]string) {
	qu.QueryAndMap(func(buf []string) {
		result = append(result, buf)
	}, param...)
	return
}

//按照给定的参数查询sql并对每一行数据执行mapper
func (qu *QueryUtil) QueryAndMap(mapper func([]string), param ...interface{}) {
	//进行查询
	rows, err := qu.stmt.Query(param...)
	PanicIfErr(err)

	//从结果集中取每一行读取到buf中，然后调用f函数进行处理
	for rows.Next() {
		//its中的元素均为buf元素的指针，数据可直接从buf中取到
		err = rows.Scan(qu.iBuf...)
		PanicIfErr(err)
		//取buf中的值放入vBuf作为f函数的参数，避免传指针可能出现的问题
		var vBuf []string
		for _, bytes := range qu.buf {
			vBuf = append(vBuf, string(*bytes))
		}
		mapper(vBuf)
	}
}

const selectStr = "select"
const fromStr = "from"
const asStr = "as"

//读取sql文件中的字段别名或字段名
func readColumnNames(sqlString string) []string {
	var columnNames []string
	start := strings.Index(strings.ToLower(sqlString), selectStr) + len(selectStr)
	end := strings.Index(strings.ToLower(sqlString), fromStr)
	columns := strings.Split(strings.TrimSpace(sqlString[start:end]), ",")
	if len(columns) == 1 && columns[0] == "," {
		panic("can not found column in sql [" + sqlString + "]")
	}
	for _, column := range columns {
		//定位as的位置，如果存在as取别名作为字段名
		column = strings.ReplaceAll(column, "\r", "")
		column = strings.ReplaceAll(column, "\n", "")
		column = strings.ReplaceAll(column, "`", "")
		asIndex := strings.Index(strings.ToLower(column), asStr)
		if asIndex != -1 {
			columnAlias := strings.Trim(column[asIndex+len(asStr):], "' ")
			columnNames = append(columnNames, columnAlias)
		} else {
			columnNames = append(columnNames, strings.Trim(column, "' "))
		}
	}
	return columnNames
}
func (qu *QueryUtil) GetColumnNames() []string {
	return qu.columnNames
}
