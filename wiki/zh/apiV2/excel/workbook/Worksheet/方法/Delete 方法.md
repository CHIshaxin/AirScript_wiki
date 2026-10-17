---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DFRo4cpr
---

# Delete (方法)

删除对象。

## 说明

在删除 **Worksheet** 时，此方法默认情况下显示一个对话框，用于提示用户确认是否删除。

## 示例

```javascript
/*本示例删除当前工作表。*/
function test() {
    ActiveSheet.Delete()
}
```

```javascript
/*本示例判断如果工作表数量大于1，则删除第一张工作表。*/
function test() {
    if (Worksheets.Count  > 1) {
        Worksheets.Item(1).Delete()
    }
}
```
