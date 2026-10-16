---
title: "DisplayBlanksAs (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/DisplayBlanksAs%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: NMLpOyae
---

# DisplayBlanksAs (属性)

返回或设置在图表中绘制空白单元格的方式。可以是 **XlDisplayBlanksAs** 常量之一。**Long** 类型，可读写。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 中不绘制空白单元格。*/
function test() {
    Application.Charts.Item("Chart1").DisplayBlanksAs = xlNotPlotted
}
```

```javascript
/*本示例判断图表工作表 Chart2 绘制空白单元格的方式是否为将值内插入图表。*/
function test() {
    let chart = Application.Charts.Item("Chart2")
    console.log(chart.DisplayBlanksAs == xlInterpolated)
}
```
