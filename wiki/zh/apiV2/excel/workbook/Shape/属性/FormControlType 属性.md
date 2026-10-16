---
title: "FormControlType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/FormControlType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B_q-DCNL
---

# FormControlType (属性)

返回 ET 控件 （ET 控件：ET 本身具有的控件，而不是 ActiveX 控件。） **[XlFormControl](./../../枚举/XlFormControl 枚举.html)** 类型，只读。

## 说明

不能对 ActiveX 控件使用该属性（ **[Shape](./../../Shape/Shape 对象.html)** 对象的 **Type** 属性必须返回 **msoFormControl**）。

## 示例

```javascript
/*本示例清除第一张工作表中的所有ET复选框。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        if (shapes.Item(i).Type == msoFormControl) {
            if (shapes.Item(i).FormControlType == xlCheckBox) {
                shapes.Item(i).ControlFormat.Value = false
            }
        }
    }
}
```

```javascript
/*本示例判断活动工作表中第二个形状是否为ET复选框。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    if (shapes.Item(2).FormControlType == xlCheckBox) {
        console.log("第二个形状是复选框")
    } else {
        console.log("第二个形状不是复选框")
    }
}
```
