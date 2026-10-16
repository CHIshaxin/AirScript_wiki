---
title: "MultiSelect 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E5%B1%9E%E6%80%A7/MultiSelect%20%E5%B1%9E%E6%80%A7.html
---
# MultiSelect (属性)

返回或设置指定列表框的选择模式。可为以下常量之一：**xlNone**、**xlSimple** 或 **xlExtended**。**Long** 类型，可读写。

## 说明

如果为单项选择 (**xlNone**)，则每次只允许选择一项。单击鼠标或按空格键将取消原先的选择项，并选定正在单击的项。

如果为简单的多项选择 (**xlSimple**)，当用鼠标单击选定某选项，或在突出显示选项上按空格键时，可切换该选项的选定状态，此方式适用于经常需要选定多个选项的列表。

如果为扩展多项选择 (**xlExtended**)，其作用与单项选择列表框一样，单击某选项将选定该选项并取消选定其他选项。按住 Shift 单击鼠标或按箭头键，可从当前选项开始顺序选定多个选项。按住 Ctrl 单击选项，可在选定项集合中添加选中项。此方式适用于允许多项选择但又不经常使用的列表。

可使用 **Value** 属性或 **ListIndex** 属性返回或设置单项选择列表框的选定项。

不能使用 **LinkedCell** 属性对多项选择列表框进行链接。

## 示例

```javascript
/*本示例创建一个简单多项选择列表框。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.AddFormControl(xlListBox, 10, 10, 100, 100)
    shape.ControlFormat.MultiSelect = xlSimple
}
```

```javascript
/*本示例显示活动工作表的第二个形状（列表框）的选择模式是否为单项选择。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.ControlFormat.MultiSelect == xlNone)
}
```
