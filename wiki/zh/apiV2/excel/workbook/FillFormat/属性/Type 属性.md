---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B34q0aYV
---

# Type (属性)

返回代表填充类型的 **MsoFillType** 值。

## 说明

此属性也可返回 **xlAutomatic** 或 **xlNone**。

在 ET 中不使用 **msoFillBackground**常量。

## 示例

```javascript
/*本示例判断活动工作表中第二个形状的填充类型是否是纹理填充。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.Item(2).Fill
    if (fillFormat.Type == msoFillTextured) {
        console.log("是纹理填充")
    } else {
        console.log("不是纹理填充")
    }
}
```

```javascript
/*本示例显示第一张工作表中第一个形状的填充类型是否是实心填充。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(1).Fill
    console.log(fillFormat.Type == msoFillSolid)
}
```
