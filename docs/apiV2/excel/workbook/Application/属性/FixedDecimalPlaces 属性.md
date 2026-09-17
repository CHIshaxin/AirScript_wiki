---
title: "FixedDecimalPlaces 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/FixedDecimalPlaces%20%E5%B1%9E%E6%80%A7.html
---
# FixedDecimalPlaces (属性)

返回或设置当 **FixedDecimal** 属性设置为 **True** 时使用的固定小数位数。**Long** 类型，可读写。

## 示例

```javascript
/*本示例将 FixedDecimal 属性设置为true，并将 FixedDecimalPlaces 属性设为 4。本示例运行后，在工作表中输入“30000”将生成“3”，输入“12500”生成“1.25”。*/
function test() {
    Application.FixedDecimal = true
    Application.FixedDecimalPlaces = 4
}
```

```javascript
/*本示例演示将FixedDecimalPlaces属性的值设置为3，并显示属性的值。*/
function test() {
    Application.FixedDecimalPlaces = 3
    console.log("FixedDecimalPlaces属性的值为：" + Application.FixedDecimalPlaces)
}
```
