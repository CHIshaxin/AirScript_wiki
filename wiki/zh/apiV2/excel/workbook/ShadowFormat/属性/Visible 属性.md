---
title: "Visible (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BFwb3WyU
---

# Visible (属性)

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*本示例为第一张工作表中的第三个形状设置阴影可见并设置阴影的格式类型和水平偏移量。阴影相对形状向右偏移 6 磅。*/
function test() {
    let shadow = Application.Worksheets.Item(1).Shapes.Item(3).Shadow
    shadow.Visible = true
    shadow.Type = msoShadow18
    shadow.OffsetX = 6
}
```

```javascript
/*本示例显示活动工作表中第一个形状阴影是否可见。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(1).Shadow
    console.log(shadow.Visible == msoTrue)
}
```
