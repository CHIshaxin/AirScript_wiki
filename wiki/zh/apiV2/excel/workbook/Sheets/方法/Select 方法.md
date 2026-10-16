---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sheets/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DsPyXlkk
---

# Select (方法)

选择对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Replace | any | 可选 | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 示例

```javascript
/*此示例选中活动工作簿中名称为“Sheet1”和“Sheet2”的工作表。*/
function test() {
    let names = ["Sheet1", "Sheet2"]
    Application.ActiveWorkbook.Sheets.Item(names).Select()
}
```

```javascript
/*此示例选中所有的工作表。*/
function test() {
    Sheets.Select()
}
```
