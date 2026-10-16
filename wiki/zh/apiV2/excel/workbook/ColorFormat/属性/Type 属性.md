---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorFormat/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DYlE1t1Q
---

# Type (属性)

返回一个代表颜色格式类型的 **MsoColorType** 值。

## 示例

```javascript
/*本示例判断活动工作表中第一个形状前景填充色的颜色类型，并显示结果。*/
function test() {
    if (ActiveSheet.Shapes.Item(1).Fill.ForeColor.Type == msoColorTypeScheme) {
        console.log("颜色由应用程序特定的方案定义。")
    } else {
        console.log("其他类型")
    }
}
```

```javascript
/*本示例显示第一张工作表上第二个形状背景填充色的颜色类型是否为 msoColorTypeRGB。*/
function test() {
    let colorFormat = Worksheets.Item(1).Shapes.Item(2).Fill.BackColor
    console.log(colorFormat.Type == msoColorTypeRGB)
}
```
