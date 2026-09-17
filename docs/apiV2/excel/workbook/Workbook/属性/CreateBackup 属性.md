---
title: "CreateBackup 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/CreateBackup%20%E5%B1%9E%E6%80%A7.html
---
# CreateBackup (属性)

如果保存文件时创建备份文件，则该属性值为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例检查在保存当前工作簿时是否创建备份文件。如果创建，则显示一条消息。*/
function test() {
    if (ActiveWorkbook.CreateBackup) {
        console.log("Remember, there is a backup copy of this workbook")
    }
}
```

```javascript
/*本示例演示CreateBackup属性的数据类型。*/
function test() {
    console.log(typeof ActiveWorkbook.CreateBackup)
}
```
