---
title: "Location 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/Location%20%E5%B1%9E%E6%80%A7.html
---
# Location (属性)

获取或设置一个 **String** 类型的值，该值代表指定 **PivotTable** 主体中左上角的单元格地址。可读/写。

## 示例

```javascript
/*本示例将活动工作表中第一张数据透视表左上角移动到 A1 单元格*/
function test() {
    ActiveSheet.PivotTables(1).Location = "A1"
}
```

```javascript
/*本示例将数据透视表左上角移动到 C30 单元格*/
function test() {
    ActiveSheet.PivotTables(1).Location = Range("C30").Address()
}
```
