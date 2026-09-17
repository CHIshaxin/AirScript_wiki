---
title: "Next 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Next%20%E5%B1%9E%E6%80%A7.html
---
# Next (属性)

返回代表下一个工作表的 Worksheet对象。

## 说明

如果该对象为区域，则此属性会模拟 Tab 键，但会返回下一个单元格，而不选中它。

在处于保护状态的工作表中，此属性返回下一个未锁定单元格。在未保护的工作表中，此属性总是返回紧靠指定单元格右边的单元格。

## 示例

```javascript
/*此示例选定 sheet1 中下一个未锁定单元格。如果 sheet1 未保护，选定的单元格将是紧靠活动单元格右边的单元格。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Next.Select()
}
```

```javascript
/*此示例判断如果工作表数量小于2，则新增一张工作表，通知用户新增工作表的下一个工作表的名称。*/
function test() {
    if (Worksheets.Count < 2) {
        let ws = Worksheets.Add()
        console.log(ws.Next.Name)
    }
}
```
