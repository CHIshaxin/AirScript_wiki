---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Styles/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: eYHY2P1H
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*此示例通过设置“常规”样式的 Bold 属性来更改活动工作簿中的该样式。*/
function test() {
    Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true
}
```

```javascript
/*本示例显示活动工作簿第一个样式是否包含合并的单元格。*/
function test() {
    console.log(ActiveWorkbook.Styles.Item(1).MergeCells)
}
```
