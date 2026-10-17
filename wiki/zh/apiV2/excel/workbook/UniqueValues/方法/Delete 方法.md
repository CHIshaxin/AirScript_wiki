---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UniqueValues/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Bx0Oy89N
---

# Delete (方法)

删除指定的条件格式规则对象。

## 示例

```javascript
/*本示例删除工作表 Sheet1 上区域 C1:C10 中第二个（UniqueValues）条件格式。*/
function test() {
    let uniqueValues = Application.Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(2)
    uniqueValues.Delete()
}
```

```javascript
/*本示例判断如果活动工作表上区域 C1:C10 中第一个（UniqueValues）条件格式应用到的单元格区域的地址为“$C$5:$C$10”，则删除该条件格式。*/
function test() {
    let uniqueValues = ActiveSheet.Range("C1:C10").FormatConditions.Item(1)
    if (uniqueValues.AppliesTo.Address() == "$C$5:$C$10") {
        uniqueValues.Delete()
    }
}
```
