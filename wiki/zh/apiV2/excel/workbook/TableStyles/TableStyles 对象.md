---
title: "TableStyles (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyles/TableStyles%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CpqXymT0
---

# TableStyles (对象)

代表可应用于表格的样式。

## 说明

表格样式提供了一种为整个表格或数据透视图设置格式的方式。表格样式取代了用于为整个表格设置格式的现有的自动套用格式功能。

表格样式与自动套用格式在以下几个方面不同：

- 可以创建和重用自定义表格样式。
- 表格样式可处理主题。
- 如果更改文档主题的配色方案和/或字体方案，则将更改内置表格样式的外观。
- 当对象发生变化时，表格样式可以将样式重新应用于像数据透视表和表格之类的对象。该表格将记住应用于对象的样式，在添加、删除、隐藏和显示单元格时，表格将相应地进行重新显示。
- 表格样式在功能区中具有可见的用户界面。

## 示例

```javascript
/*本示例显示第一张工作簿中所使用的样式数量。*/
function test() {
    console.log(Workbooks.Item(1).TableStyles.Count)
}
```

```javascript
/*本示例向活动工作簿的表格样式中添加“样式1”。*/
function test() {
    ActiveWorkbook.TableStyles.Add("样式1")
}
```
