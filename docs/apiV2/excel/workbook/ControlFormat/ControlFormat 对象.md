---
title: "ControlFormat 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/ControlFormat%20%E5%AF%B9%E8%B1%A1.html
---
# ControlFormat (对象)

包含 ET 控件 （ET 控件：ET 本身具有的控件，而不是 ActiveX 控件。）属性。

## 示例

```javascript
/*下例为第一张工作表上的列表框控件设置填充区域。*/
function test() {
    Worksheets.Item(1).Shapes.Item(1).ControlFormat.ListFillRange = "A1:A10"
}
```

```javascript
/*本示例删除活动工作表上第二个形状（组合框）的所有数据项。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.ControlFormat.RemoveAllItems()
}
```
