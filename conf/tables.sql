SELECT
    TABLE_NAME as '表名',
    TABLE_COMMENT as '描述',
    ENGINE as '数据库引擎'
FROM
    information_schema.`TABLES`
WHERE
    TABLE_SCHEMA = ?