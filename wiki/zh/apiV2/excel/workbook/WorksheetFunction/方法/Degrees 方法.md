---
title: "Degrees (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Degrees%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: VUPu8K0o
---

# Degrees (方法)

将弧度转换为度。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Angle - 要转换的以弧度表示的角度。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B1等单元格分别赋值，并使用Degrees方法将弧度转换为度。*/
function test() {
    Range("B1").Value2 = 60
    Range("B2").Value2 = 30
    Range("B3").Value2 = 10.5
    let degrees1 = Application.WorksheetFunction.Degrees(Range("B1").Value2)
    let degrees2 = Application.WorksheetFunction.Degrees(Range("B2").Value2)
    let degrees3 = Application.WorksheetFunction.Degrees(Range("B3").Value2)
    console.log(degrees1)
    console.log(degrees2)
    console.log(degrees3)
}
```

```javascript
/*本示例使用Degrees方法将弧度转换为度，并将结果赋值于D3等单元格。*/
function test() {
    Range("D3").Value2 = Application.WorksheetFunction.Degrees(-100)
    Range("D4").Value2 = Application.WorksheetFunction.Degrees(87)
    Range("D5").Value2 = Application.WorksheetFunction.Degrees(43.5)
}
```
