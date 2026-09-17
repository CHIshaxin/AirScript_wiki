---
title: "AutoFilter 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/AutoFilter%20%E5%B1%9E%E6%80%A7.html
---
# AutoFilter (属性)

如果筛选已打开，则返回一个 **[AutoFilter](./../../AutoFilter/AutoFilter 对象.html)** 对象。只读。

## 说明

如果筛选已关闭，该属性返回 **null**。

要为工作表创建 **[AutoFilter](./../../AutoFilter/AutoFilter 对象.html)** 对象，必须手动打开或使用 **[Range](./../../Range/Range 对象.html)** 对象的 **AutoFilter** 方法打开工作表上某个区域上的自动筛选功能。

## 示例

```javascript
/*下面的示例返回当前工作表的自动筛选。*/
function test() {
    let returnValue = Worksheets.Item(1).AutoFilter
}
```

```javascript
/*本示例显示活动工作表中筛选器的总数，以及活动工作表是否处于自动筛选筛选器模式。*/
function test() {
    console.log(`活动工作表中筛选器的总数：${ActiveSheet.AutoFilter.Filters.Count}\n活动工作表是否处于自动筛选筛选器模式：${ActiveSheet.AutoFilter.FilterMode}`)
}
```
