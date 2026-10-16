---
title: "NumberFormatLocal (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/NumberFormatLocal%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: F-4CEBnc
---

# NumberFormatLocal (属性)

以采用用户语言字符串的形式返回或设置一个 **String** 值，它代表对象的格式代码。

## 说明

**Format** 函数使用的格式代码字符串与 **NumberFormat**和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

## 示例

```javascript
/*本示例显示活动工作簿第一个样式的数字格式。*/
function test() {
    console.log("活动工作簿第一个样式的数字格式为" + ActiveWorkbook.Styles.Item(1).NumberFormatLocal)
}
```

```javascript
/*本示例为单元格A1赋值，并将A1单元格的数字格式设置为百分比格式。*/
function test() {
    Worksheets("Sheet1").Range("A1").Value2 = 0.3
    Worksheets("Sheet1").Range("A1").NumberFormatLocal = "0.00%"
}
```
