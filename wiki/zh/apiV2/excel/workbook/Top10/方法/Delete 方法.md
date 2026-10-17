---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Top10/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CjWtXyx9
---

# Delete (方法)

删除指定的条件格式规则对象。

## 示例

```javascript
/*本示例删除活动工作表上区域 B1:B10 中第一个（Top10）条件格式。*/
function test() {
    let top = ActiveSheet.Range("B1:B10").FormatConditions.Item(1)
    top.Delete()
}
```

```javascript
/*本示例判断如果活动工作表上区域 B1:B10 中第二个（Top10）条件格式的TopBottom属性为xlTop10Top，则删除该条件格式。*/
function test() {
    let top = ActiveSheet.Range("B1:B10").FormatConditions.Item(2)
    if (top.TopBottom == xlTop10Top) {
        top.Delete()
    }
}
```
