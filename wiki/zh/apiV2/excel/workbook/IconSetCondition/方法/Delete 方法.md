---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSetCondition/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BiCy-czR
---

# Delete (方法)

删除指定的条件格式规则对象。

## 示例

```javascript
/*本示例删除第一张工作表上区域 A1:A10 中第一个图标集条件格式。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1)
    iconSet.Delete()
}
```

```javascript
/*本示例判断如果活动工作表上单元格区域 A1:A10 的第二个图标集条件格式仅展示图标，则删除该条件格式。*/
function test() {
    let iconSet = ActiveSheet.Range("A1:A10").FormatConditions.Item(2)
    if (iconSet.ShowIconOnly) {
        iconSet.Delete()
    }
}
```
