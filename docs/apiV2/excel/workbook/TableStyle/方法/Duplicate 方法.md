---
title: "Duplicate 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/%E6%96%B9%E6%B3%95/Duplicate%20%E6%96%B9%E6%B3%95.html
---
# Duplicate (方法)

复制 **[TableStyle](./../../TableStyle/TableStyle 对象.html)** 对象，并返回对新复制对象的引用。

## 说明

如果未指定名称，则 **Duplicate** 方法将使用与用户界面默认名称相同的默认名称。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| NewTableStyleName | any | 可选 | 新表样式的名称。 |

## 返回值

TableStyle

## 示例

```javascript
/*本示例复制活动工作簿上的表格样式中的最后一个样式。*/
function test() {
    let tableStyles = ActiveWorkbook.TableStyles
    tableStyles.Item(tableStyles.Count).Duplicate("样式5")
}
```

```javascript
/*本示例复制第一张工作簿中指定索引的表格样式，并显示复制样式的名称。*/
function test() {
    let tableStyle = Workbooks.Item(1).TableStyles.Item(1).Duplicate("样式三")
    console.log(tableStyle.Name)
}
```
