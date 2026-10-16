---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BViJbvrp
---

# Delete (方法)

删除指定的条件格式规则对象。

## 示例

```javascript
/*本示例删除工作表 Sheet1 上区域 A1:A10 中第一个为数据条的条件格式。*/
function test() {
    let databar = Application.Worksheets.Item("Sheet1").Range("A1:A10").FormatConditions.Item(1)
    databar.Delete()
}
```

```javascript
/*本示例判断如果活动工作表上区域 A1:A10 中第一个为数据条的条件格式的填充色为纯色，则删除该数据条。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    if (databar.BarFillType == xlDataBarFillSolid) {
        databar.Delete()
    }
}
```
