---
title: "Visible (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DC6lP72P
---

# Visible (属性)

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*本示例显示活动工作簿中索引为1的名称是否可见。*/
function test() {
    console.log(ActiveWorkbook.Names.Item(1).Visible)
}
```

```javascript
/*本示例将 Sheet1 中名为“test1”的名称设置为不可见。*/
function test() {
    Application.Worksheets.Item("Sheet1").Names.Item("test1").Visible = false
}
```
