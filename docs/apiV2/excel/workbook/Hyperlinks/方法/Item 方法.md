---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlinks/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*下例激活 E5 单元格的第一个超链接。*/
function test() {
    Application.Worksheets.Item(1).Range("E5").Hyperlinks.Item(1).Follow()
}
```

```javascript
/*本示例取消工作表Sheet1中的第一个超链接。*/
function test() {
    Application.Worksheets.Item("Sheet1").Hyperlinks.Item(1).Delete()
}
```
