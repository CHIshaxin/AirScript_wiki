---
title: "IconSet 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSetCondition/%E5%B1%9E%E6%80%A7/IconSet%20%E5%B1%9E%E6%80%A7.html
---
# IconSet (属性)

返回或设置一个 **[IconSets](./../../IconSets/IconSets 对象.html)** 集合，该集合指定用于条件格式的图标集。

## 说明

通过使用 **[Workbook](./../../Workbook/Workbook 对象.html)** 对象的 **[IconSets](./../../IconSets/IconSets 对象.html)** 属性，您可以指定图标集。例如，`Selection.FormatConditions(1).IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)` 将对条件格式应用由三个交通信号灯图标组成的图标集。

## 示例

```javascript
/*本示例显示活动工作表上区域 B1:B10 中第二个图标集条件格式使用的图标集的数量。*/
function test() {
    let iconSet = ActiveSheet.Range("B1:B10").FormatConditions.Item(2)
    console.log(iconSet.IconSet.Count)
}
```

```javascript
/*本示例显示第一张工作表上区域 B1:B10 中第一个图标集条件格式使用的图标集的名称是否为xl3Flags。*/
function test() {
    let iconSet = Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(1)
    console.log(iconSet.IconSet.ID == xl3Flags)
}
```
