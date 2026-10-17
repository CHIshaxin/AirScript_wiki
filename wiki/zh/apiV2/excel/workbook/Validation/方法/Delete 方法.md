---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cd3xK-B0
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例删除单元格 A1 的数据有效性验证。*/
function test() {
    Range("A1").Validation.Delete()
}
```

```javascript
/*本示例判断如果单元格 A1 的数据有效性验证的警告样式为xlValidAlertWarning，则删除该数据有效性验证。*/
function test() {
    let validation = ActiveSheet.Range("A1").Validation
    if (validation.AlertStyle == xlValidAlertWarning) {
        validation.Delete()
    }
}
```
