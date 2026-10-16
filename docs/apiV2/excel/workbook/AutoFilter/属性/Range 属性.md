---
title: "Range 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AutoFilter/%E5%B1%9E%E6%80%A7/Range%20%E5%B1%9E%E6%80%A7.html
---
# Range (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表应用指定的“自动筛选”的区域。

## 示例

```javascript
/*下例将应用第一张工作表的“自动筛选”的地址保存在变量中。*/
function test() {
    let address = Application.Worksheets.Item(1).AutoFilter.Range.Address()
}
```

```javascript
/*下例选中活动工作表自动筛选应用到的区域。*/
function test() {
    ActiveSheet.AutoFilter.Range.Select()
}
```
