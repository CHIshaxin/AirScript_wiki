---
title: "UsedRange 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/UsedRange%20%E5%B1%9E%E6%80%A7.html
---
# UsedRange (属性)

返回一个 Range 对象，该对象表示指定工作表上所使用的区域。只读。

## 示例

```javascript
/*本示例选定 Sheet1 中的已用区域。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveSheet.UsedRange.Select()
}
```

```javascript
/*本示例显示活动工作表总共使用多少单元格。*/
function test() {
    console.log("A total of " + ActiveSheet.UsedRange.Count + "  cells are used ")
}
```
