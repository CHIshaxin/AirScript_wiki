---
title: "Acos 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Acos%20%E6%96%B9%E6%B3%95.html
---
# Acos (方法)

返回数字的反余弦值。反余弦值是余弦值为 *Arg1* 的角度。返回的角度以弧度给出，范围从 0（零）到 pi。

## 说明

如果要将结果从弧度转换为度，则请将结果乘以 180/PI()，或者使用 Degrees 方法。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 所需角度的余弦值，必须介于 -1 到 1 之间。 |

## 示例

```javascript
/*本示例计算数字的反余弦值，并将结果分别赋值于A2等单元格。*/
function test() {
    Range("A2").Value2 = Application.WorksheetFunction.Acos(-1)
    Range("B2").Value2 = Application.WorksheetFunction.Acos(0.93969262)
    Range("C2").Value2 = Application.WorksheetFunction.Acos(1)
    Range("D2").Value2 = Application.WorksheetFunction.Acos(-0.08715574)
    Range("E2").Value2 = Application.WorksheetFunction.Acos(0)
}
```

```javascript
/*本示例为B3等单元格分别赋值，再将其用于计算反余弦值。*/
function test() {
    Range("B3").Value2 = 0.81915204
    Range("C3").Value2 = -0.25881905
    Range("D3").Value2 = 0.5
    let acos1 = Application.WorksheetFunction.Acos(Range("B3").Value2)
    let acos2 = Application.WorksheetFunction.Acos(Range("C3").Value2)
    let acos3 = Application.WorksheetFunction.Acos(Range("D3").Value2)
    console.log(acos1)
    console.log(acos2)
    console.log(acos3)
}
```
