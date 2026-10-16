---
title: "Weight (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Border/%E5%B1%9E%E6%80%A7/Weight%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: dGLocRwo
---

# Weight (属性)

返回或设置一个 **XlBorderWeight**值，它代表边框的粗细。

## 示例

```javascript
/*本示例对 Sheet1 上椭圆一的边框粗细进行设置。*/
function test() {
    Application.Worksheets.Item("Sheet1").Ovals(1).Border.Weight = xlMedium
}
```

```javascript
/*本示例显示 I1 单元格左边缘边框的粗细是否为 xlThin 。*/
function test() {
    console.log(Range("I1").Borders.Item(xlEdgeLeft).Weight == xlThin)
}
```
