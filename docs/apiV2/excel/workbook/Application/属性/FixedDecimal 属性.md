---
title: "FixedDecimal 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/FixedDecimal%20%E5%B1%9E%E6%80%A7.html
---
# FixedDecimal (属性)

如果该属性被设置为 **True**，则以后输入的所有数据都根据 **FixedDecimalPlaces** 属性所设置的固定小数位数设置格式。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例将 FixedDecimal 属性设置为true，并将 FixedDecimalPlaces 属性设为 4。本示例运行后，在工作表中输入“30000”将生成“3”，输入“12500”生成“1.25”。*/
function test() {
    Application.FixedDecimal = true
    Application.FixedDecimalPlaces = 4
}
```

```javascript
/*本示例演示 FixedDecimal 属性的数据类型，并将 FixedDecimal 属性的值设置为false。*/
function test() {
    console.log("FixedDecimal属性的数据类型为：" + typeof Application.FixedDecimal)
    Application.FixedDecimal = false
}
```
