---
title: "Type 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
---
# Type (属性)

返回一个代表形状类型的 **MsoShapeType** 值。

## 示例

```javascript
/*本示例先在第一张工作表中创建shpRange对象，然后判断对象类型是否为图表并显示。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([4, 3])
    if (shpRange.Type == msoChart) {
        console.log("这是图表")
    } else {
        console.log("这不是图表")
    }
}
```

```javascript
/*本示例先在活动工作表中创建shpRange对象，然后显示对象类型是否为图片。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([1])
    console.log(shpRange.Type == msoPicture)
}
```
