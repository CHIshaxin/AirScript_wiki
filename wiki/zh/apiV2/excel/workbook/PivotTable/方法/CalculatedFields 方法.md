---
title: "CalculatedFields (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/CalculatedFields%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: aH8zU6dt
---

# CalculatedFields (方法)

返回一个 **[CalculatedFields](./../../CalculatedFields/CalculatedFields 对象.html)** 集合，该集合表示指定数据透视表中的所有计算字段。只读。

## 返回值

CalculatedFields

## 示例

```javascript
/*本示例使计算结果字段不能被拖至行。*/
function test() {
    for (let i = 1; i <= Worksheets.Item(1).PivotTables("Pivot1").CalculatedFields().Count; i++) {
        Worksheets.Item(1).PivotTables("Pivot1").CalculatedFields().Item(i).DragToRow = false
    }
}
```

```javascript
/*本示例在数据透视表 Pivot1 中添加计算字段，并显示所有计算字段的数量*/
function test() {
    ActiveSheet.PivotTables("Pivot1").CalculatedFields().Add("y*z", "=y*z")
    console.log("数据透视表Pivot1中计算字段的数量为" + ActiveSheet.PivotTables("Pivot1").CalculatedFields().Count)
}
```
