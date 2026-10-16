---
title: "MatchCase (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sort/%E5%B1%9E%E6%80%A7/MatchCase%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DjeVnTFi
---

# MatchCase (属性)

如果设置为 **True**，则执行区分大小写的排序，如果设置为 **False**，则执行不区分大小写的排序。可读/写。

## 示例

```javascript
/*本示例显示活动工作表排序时是否区分大小写。*/
function test() {
    console.log(Application.ActiveSheet.Sort.MatchCase)
}
```

```javascript
/*本示例设置第一张工作表排序区分大小写，并通知用户。*/
function test() {
    Application.Sheets.Item(1).Sort.MatchCase = true
    console.log("执行排序区分大小写")
}
```
