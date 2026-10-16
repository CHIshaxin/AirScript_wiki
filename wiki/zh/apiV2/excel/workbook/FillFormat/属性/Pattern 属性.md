---
title: "Pattern (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/Pattern%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: kPK8w-7d
---

# Pattern (属性)

返回一个代表填充图案的 **MsoPatternType** 值。只读。

## 说明

使用 **Patterned** 方法可设置填充的图案类型。

## 示例

```javascript
/*本示例向活动工作表中添加一个矩形，并显示填充图案是否是混和图案。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill
    console.log(fillFormat.Pattern == msoPatternMixed)
}
```

```javascript
/*本示例判断第一张工作表中第一个形状的填充图案。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    if (fillFormat.Pattern == msoPatternNarrowVertical) {
        console.log("窄竖线图案")
    } else {
        console.log("其它图案")
    }
}
```
