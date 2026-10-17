---
title: "Activate (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/Activate%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B2vf34M0
---

# Activate (方法)

使当前工作表成为活动工作表。

## 说明

调用此方法等同于单击工作表的标签。

## 示例

```javascript
/*此示例激活工作表 Sheet1。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
}
```

```javascript
/*本示例激活第二张工作表*/
function test() {
    Application.Worksheets.Item(2).Activate()
}
```
