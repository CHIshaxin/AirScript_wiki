---
title: "Strikethrough (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/Strikethrough%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BoiqdukQ
---

# Strikethrough (属性)

如果文字中间有一条水平删除线，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例将 Sheet1 中活动单元格的字体设置为加删除线*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Font.Strikethrough = true
}
```

```javascript
/*本示例判断第一个工作表中单元格D5的字体如果有删除线，则通知用户，并取消删除线*/
function test() {
    let font = Application.Worksheets.Item(1).Range("D5").Font
    if (font.Strikethrough) {
        console.log("取消字体的删除线")
        font.Strikethrough = false
    }
}
```

``
