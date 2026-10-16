---
title: "ThreeD (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartFormat/%E5%B1%9E%E6%80%A7/ThreeD%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C53hk6GM
---

# ThreeD (属性)

返回一个 **[ThreeDFormat](./../../ThreeDFormat/ThreeDFormat 对象.html)** 对象，该对象包含指定图表的三维效果格式属性。只读。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 图表区的底部凹凸效果类型是否为 Angle 凹凸效果。*/
function test() {
    let chartformat = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea.Format
    console.log(chartformat.ThreeD.BevelBottomType == msoBevelAngle)
}
```

```javascript
/*此示例设置工作表 Sheet1 第一个内嵌图表区延伸表面为线框材料。*/
function test() {
    let chartformat = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format
    chartformat.ThreeD.PresetMaterial = msoMaterialWireFrame
}
```
