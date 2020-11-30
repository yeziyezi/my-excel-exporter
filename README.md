# my-excel-exporter
## todo list
##### BUG
- 当SheetName长度超过31会被excel自动截断，如果两个Sheet截断后SheetName相同会导致一些问题。解决方案：增加SheetName与tableName的映射关系，用数字代替tableName作为SheetName
##### <del>todo 调整表格输出的样式</del>
- <del>根据字段长度调整行宽</del>
- <del>首个sheet增加跳转到其他sheet的链接</del>
- <del>增加其他sheet跳回首页的超链接</del>
- <del>给超链接增加蓝色带下划线的style</del>

##### todo 优化代码
- 将excel的某些方法的参数改为直接传入Config对象
- 将TableListSheetExcelUtil合并到ExcelUtil里
##### todo 报错提示友好化
##### todo 完善注释
##### todo 完善readme
##### todo 打zip包作为release
## 参考资料
- **excelize官方文档** https://xuri.me/excelize/zh-hans/
- **go-sql-driver** https://github.com/go-sql-driver/mysq