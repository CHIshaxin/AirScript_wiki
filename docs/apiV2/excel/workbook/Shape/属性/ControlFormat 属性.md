---
title: "ControlFormat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/ControlFormat%20%E5%B1%9E%E6%80%A7.html
---
# ControlFormat (属性)

返回一个 **[ControlFormat](./../../ControlFormat/ControlFormat 对象.html)** 对象，该对象包含 ET 控件 （ET 控件：ET 本身具有的控件，而不是 ActiveX 控件。）属性。只读。

## 示例

```javascript
/*本示例删除第一张工作表中列表框中选定的数据项。如果 Shapes(2) 代表的不是列表框，则本示例将无效。*/
function test() {
    let controlFormat = Worksheets.Item(1).Shapes.Item(2).ControlFormat
    controlFormat.RemoveItem(controlFormat.ListIndex)
}
```

```javascript
/*本示例判断如果活动工作表中第三个形状是复选框，则不勾选复选框。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(3).Type == msoFormControl && shapes.Item(3).FormControlType == xlCheckBox) {
        shapes.Item(3).ControlFormat.Value = false
    }
}
```
