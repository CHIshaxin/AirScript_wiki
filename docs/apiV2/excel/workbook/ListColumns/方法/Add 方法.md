---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListColumns/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
---
# Add (方法)

向列表对象中添加新列。

## 说明

如果不指定 *Position*，就在最右边添加一列，并自动为该列生成名称。该列的名称可在添加后更改。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Position | any | 可选 | Integer 类型。从 1 开始指定新列的相对位置。以前位于此位置的列则向后移动。 |

## 返回值

一个代表新列的 ListColumn 对象。

## 示例

```javascript
/*下例给工作簿的第一张工作表的默认 ListObject 对象添加一个新列。由于未指定位置，因此在最右边添加一个新列。*/
function test() {
    let newColumn = Application.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Add()
}
```

```javascript
/*下例给活动工作表的第一个 ListObject 对象添加一个新的列，并加到列表中现有列的第二个位置。*/
function test() {
    ActiveSheet.ListObjects.Item(1).ListColumns.Add(2)
}
```
