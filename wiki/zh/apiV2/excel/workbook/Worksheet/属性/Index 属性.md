---
title: "Index (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B4FlLvtx
---

# Index (属性)

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

## 示例

```javascript
/*此示例显示由您所键入的名称指定的工作表的选项卡编号。例如，如果 Sheet4 为活动工作簿中第三个工作表选项卡，此示例就在消息框中显示“3”。*/
function test() {
    let strSheetName = "Type a sheet name, such as Sheet4."
    console.log(`This sheet is tab number ${Worksheets.Item(strSheetName).Index}`)
}
```

```javascript
/*本示例判断如果工作表数量小于等于1，则新增一张工作表，并通知用户其索引号。*/
function test() {
    let newSheet = null
    if (Worksheets.Count <= 1) {
        newSheet = Worksheets.Add()
        console.log(newSheet.Index)
    }
}
```
