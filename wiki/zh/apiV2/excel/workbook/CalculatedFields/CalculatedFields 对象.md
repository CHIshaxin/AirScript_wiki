---
title: "CalculatedFields (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CalculatedFields/CalculatedFields%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: zKs_eNed
---

# CalculatedFields (对象)

由 **[PivotField](./../PivotField/PivotField 对象.html)** 对象组成的集合，这些对象代表指定数据透视表中的所有计算字段。

## 示例

```javascript
/*此示例为活动工作表中数据透视表添加计算字段。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.CalculatedFields().Add("new date", "= date + 10")
}
```

```javascript
/*此示例显示第一张工作表中第一张数据透视表的第一个计算字段的标签文本。*/
function test() {
    let pvtField = Application.Worksheets.Item(1).PivotTables(1).CalculatedFields().Item(1)
    console.log(pvtField.Caption)
}
```
