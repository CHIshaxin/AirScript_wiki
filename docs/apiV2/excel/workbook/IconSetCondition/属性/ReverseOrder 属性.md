---
title: "ReverseOrder 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSetCondition/%E5%B1%9E%E6%80%A7/ReverseOrder%20%E5%B1%9E%E6%80%A7.html
---
# ReverseOrder (属性)

返回或设置一个 **Boolean** 值，该值指示是否保留图标集的图标次序。

## 说明

默认情况下，可以使用的大多数图标集都将使用较多的图像与较大值关联。例如，“三色交通灯”图标集将绿色圆应用于区域上限值，而将红色圆应用于区域的最低值。如果数据的较低值更有意义，如马拉松时间结果，则可能希望反转这些图标应用于区域的次序，使绿色圆与最低值关联。

如果 **[IconSetCondition](./../../IconSetCondition/IconSetCondition 对象.html)** 对象的 **IconSet** 属性为 **xlCustomSet**，**ReverseOrder** 属性将返回 **False**。此外，如果在 **[IconSetCondition](./../../IconSetCondition/IconSetCondition 对象.html)** 对象的 **IconSet** 属性为 **xlCustomSet** 时，尝试将 **ReverseOrder** 属性设置为 **True**，ET 将返回运行时错误。

## 示例

```javascript
/*本示例显示活动工作表上单元格区域 C1:C10 的第一个图标集条件格式是否保留图标集的图标次序。*/
function test() {
    let iconSet = ActiveSheet.Range("C1:C10").FormatConditions.Item(1)
    console.log(iconSet.ReverseOrder)
}
```

```javascript
/*本示例将第一张工作表上单元格区域 C1:C10 的第二个图标集条件格式设置为保留图标集的图标次序。*/
function test() {
    let iconSet = Worksheets.Item(1).Range("C1:C10").FormatConditions.Item(2)
    iconSet.ReverseOrder = true
}
```
