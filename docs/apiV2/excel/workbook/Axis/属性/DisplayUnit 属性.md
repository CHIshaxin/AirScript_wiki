---
title: "DisplayUnit 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/DisplayUnit%20%E5%B1%9E%E6%80%A7.html
---
# DisplayUnit (属性)

返回或设置数值轴的单位标签。 **[XlDisplayUnit](./../../枚举/XlDisplayUnit 枚举.html)** 、**xlCustom** 或 **xlNone** 类型，可读写。

## 说明

| **[XlDisplayUnit](./../../枚举/XlDisplayUnit 枚举.html)** 可为以下 **[XlDisplayUnit](./../../枚举/XlDisplayUnit 枚举.html)** 常量之一。 |
| --- |
| **xlHundredMillions** |
| **xlHundredThousands** |
| **xlMillions** |
| **xlTenThousands** |
| **xlThousands** |
| **xlHundreds** |
| **xlMillionMillions** |
| **xlTenMillions** |
| **xlThousandMillions** |
| 此单位标签还可为以下常量之一 |
| **xlCustom** |
| **xlNone** |

在绘制大数值图表时，使用单位标志可使刻度线标志易于阅读。例如，如果将数值轴的单位设置为百、千或百万，则可以在坐标轴的刻度标志上使用较小的数字值。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表的数值轴的显示单位设置为百。*/
function test() {
    let axes = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axes.DisplayUnit = xlHundreds
    axes.HasTitle = true
    axes.AxisTitle.Caption = "Rebate Amounts"
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴的显示单位设置为千，并将该数值轴的边框颜色设置为红色。*/
function test() {
    let axes = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axes.DisplayUnit = xlThousands
    axes.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
