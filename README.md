# my-excel-exporter
## Todo
##### BUG
- 当SheetName长度超过31会被excel自动截断，如果两个Sheet截断后SheetName相同会导致一些问题。解决方案：增加SheetName与tableName的映射关系，用数字代替tableName作为SheetName
##### 体验优化
- 点击回到首页时跳转到原来的单元格而不是A1
##### 优化代码
- <del>将TableListSheetExcelUtil合并到ExcelUtil里</del>
##### 报错提示友好化
##### 完善注释
##### 完善readme
##### 打zip包作为release
## 参考资料
- **excelize官方文档** https://xuri.me/excelize/zh-hans/
- **go-sql-driver** https://github.com/go-sql-driver/mysq