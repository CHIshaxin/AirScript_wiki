---
title: "Height (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DZK0ouVm
---

# Height (属性)

返回对象的高度（以磅为单位）。只读。

## 返回值

Double

## 示例

```javascript
/*此示例显示活动工作表第一个图表上第一个系列第二个数据点的高度。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    let point = ActiveChart.SeriesCollection(1).Points(2)
    console.log(`数据点的高度为 ${point.Height}`)
}
```

```javascript
/*此示例判断Chart1中第三个系列第五个数据点的高度如果超过150磅，则其填充颜色设置为红色，反之设置其填充颜色为绿色。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).Points(5)
    if (point.Height > 150) {
        point.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        console.log("数据点的高度超过150磅，填充颜色已设置为红色")
    } else {
        point.Format.Fill.ForeColor.RGB = RGB(0, 255, 0)
        console.log("数据点的高度未超过150磅，填充颜色已设置为绿色")
    }
}
```
