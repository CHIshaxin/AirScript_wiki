---
title: "Visible 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sheets/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
---
# Visible (属性)

返回或设置一个 **Variant** 值，它确定对象是否可见。

## 示例

```javascript
/*此示例将名称为“Sheet2”和“Sheet3”的工作表设置为不可见。*/
function test() {
    let names = ["Sheet2", "Sheet3"]
    Application.Sheets.Item(names).Visible = msoFalse
}
```

```javascript
/*此示例将活动工作簿中索引为1和2的工作表设置为可见。*/
function test() {
    let names = [1, 2]
    Application.ActiveWorkbook.Sheets.Item(names).Visible = msoTrue
}
```
