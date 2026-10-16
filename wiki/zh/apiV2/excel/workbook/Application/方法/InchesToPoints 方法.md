---
title: "InchesToPoints (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/InchesToPoints%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BYBzDbHZ
---

# InchesToPoints (方法)

将度量单位从英寸转换为磅 （磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Inches | double | 必填 | 指定要转换为磅值的英寸值。 |

## 返回值

Double

## 示例

```javascript
/*本示例将工作表 Sheet1 中的左边距设为 2.5 英寸。*/
function test() {
    Worksheets.Item("Sheet1").PageSetup.LeftMargin = Application.InchesToPoints(2.5)
}
```

```javascript
/*本示例将4英寸转换为磅值并显示。*/
function test() {
    console.log(Application.InchesToPoints(4))
}
```
