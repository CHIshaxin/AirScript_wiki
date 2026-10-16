---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comments/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象，包含在集合中的一个 **[Comment](./../../Comment/Comment 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | int4 | 必填 | 对象的索引号。 |

## 返回值

Comment

## 示例

```javascript
/*本示例隐藏第二条批注。*/
function test() {
    Application.Worksheets.Item(1).Comments.Item(2).Visible = false
}
```

```javascript
/*本示例显示活动工作表中第一条批注的内容。*/
function test() {
    console.log(Application.ActiveSheet.Comments.Item(1).Text())
}
```
