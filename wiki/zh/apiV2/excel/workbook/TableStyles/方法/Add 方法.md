---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyles/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DHEUy8Tl
---

# Add (方法)

创建新的 **[TableStyle](./../../TableStyle/TableStyle 对象.html)** 对象，并将其添加到集合中。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| TableStyleName | string | 必填 | 表样式的名称。 |

## 返回值

TableStyle

## 示例

```javascript
/*本示例向活动工作簿的表格样式中添加“样式一”。*/
function test() {
    ActiveWorkbook.TableStyles.Add("样式一")
}
```

```javascript
/*本示例向第一张工作簿的表格样式中添加“TableStyle2”样式，并显示新添加样式的名称。*/
function test() {
    let tableStyle = Workbooks.Item(1).TableStyles.Add("TableStyle2")
    console.log(tableStyle.Name)
}
```
