---
title: "DecimalSeparator 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/DecimalSeparator%20%E5%B1%9E%E6%80%A7.html
---
# DecimalSeparator (属性)

设置或返回用作小数分隔符的字符。**String** 类型，可读写。

## 示例

```javascript
/*本示例将“1,234,567.89”放置在单元格 A1 中，然后将系统的小数分隔符和千位分隔符更改为短划线。*/
function test() {
    Range("A1").Formula = "1,234,567.89"
    console.log("The system separators will now change.")
    //Define separators and apply.
    Application.DecimalSeparator = "-"
    Application.ThousandsSeparator = "-"
    Application.UseSystemSeparators = false
}
```

```javascript
/*本示例判断如果DecimalSeparator属性的值为“-”，则将DecimalSeparator属性的值设置为“/”，通知用户后，设置为“-”。*/
function test() {
    if (Application.DecimalSeparator == "-") {
        Application.DecimalSeparator = "/"
        console.log("当前小数分隔符为：" + Application.DecimalSeparator)
        Application.DecimalSeparator = "-"
    }
}
```
