---
title: "DataOption 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/%E5%B1%9E%E6%80%A7/DataOption%20%E5%B1%9E%E6%80%A7.html
---
# DataOption (属性)

指定 **[SortField](./../../SortField/SortField 对象.html)** 对象所指定的区域中的文本排序方式。可读/写 **[XlSortDataOption](./../../枚举/XlSortDataOption 枚举.html)** 类型。

## 说明

**xlSortDataOption**可以是 **xlSortNormal**或 **xlSortTextAsNumbers**。

## 示例

```javascript
/*本示例设置活动工作表将文本作为数字型数据进行排序。*/
function test() {
    Application.ActiveSheet.Sort.SortFields.Item(1).DataOption = xlSortTextAsNumbers
}
```

```javascript
/*本示例判断第一张工作表的文本排序方式。*/
function test() {
    let sort = Application.Sheets.Item(1).Sort.SortFields
    if (sort.Item(1).DataOption == xlSortNormal) {
        console.log("分别对数字和文本数据进行排序")
    } else {
        console.log("将文本作为数字型数据进行排序")
    }
}
```
