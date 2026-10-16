---
title: "RefersToRange 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/RefersToRange%20%E5%B1%9E%E6%80%A7.html
---
# RefersToRange (属性)

返回由 **[Name](./../../Name/Name 对象.html)** 对象引用的 **[Range](./../../Range/Range 对象.html)** 对象。只读。

## 说明

如果 **[Name](./../../Name/Name 对象.html)** 对象并不引用区域（例如，如果该对象引用的是一个常量或公式），则该属性无效。

要更改名称所引用的区域，请使用 **RefersTo** 属性。

## 示例

```javascript
/*本示例显示当前工作表打印区域中的行数和列数，确保在当前工作簿的活动工作表上建立打印区域。*/
function test() {
    let p = Application.Sheets.Item(Application.ActiveSheet.Name).Names.Item("Print_Area").RefersToRange.Value2
    console.log("Print_Area: " + p.length + " rows, " + p[0].length + " columns")
}
```

```javascript
/*本示例显示索引号为3的名称所引用单元格的宽度。*/
function test() {
    console.log(ActiveWorkbook.Names.Item(3).RefersToRange.Width)
}
```
