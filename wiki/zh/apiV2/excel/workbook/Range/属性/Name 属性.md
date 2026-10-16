---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: By52k1hI
---

# Name (属性)

返回或设置一个 **Variant** 值，它代表对象的名称。

## 示例

```javascript
/*本示例将 A1 单元格的 Name 属性设置为 test1 并显示*/
function test() {
    Range("A1").Name = 'test1'
    console.log(Range("A1").Name.toString())
}
```

```javascript
/*本示例将 A1 单元格的 Name 属性设置为 test1 并显示工作簿中所有名称的总数。*/
function test() {
    Range("A1").Name = 'test1'
    console.log(ActiveWorkbook.Names.Count)
}
```
