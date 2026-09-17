---
title: "Pi 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Pi%20%E6%96%B9%E6%B3%95.html
---
# Pi (方法)

返回数字 3.14159265358979，即数学常量 pi，精确到 15 位数字。

## 返回值

Double

## 示例

```javascript
/*本示例使用 Pi 方法获取数字 3.14159265358979，并赋值给E1单元格。*/
function test() {
    Range("E1").Value2 = WorksheetFunction.Pi()
}
```

```javascript
/*本示例使用 Pi 方法获取数字 3.14159265358979。*/
function test() {
    console.log(WorksheetFunction.Pi())
}
```
