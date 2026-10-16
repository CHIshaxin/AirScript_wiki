---
title: "PageFieldOrder (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/PageFieldOrder%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: h8uDq28e
---

# PageFieldOrder (属性)

返回或设置页字段在数据透视表布局上的顺序。可为以下 **[XlOrder](./../../枚举/XlOrder 枚举.html)** 常量之一：**xlDownThenOver** 或 **xlOverThenDown**。默认常量是 **xlDownThenOver**。**Long** 类型，可读写。

## 示例

```javascript
/*本示例使数据透视表先在一行中显示三个页字段，然后再显示下一个新行。*/
function test() {
    Application.Worksheets.Item(1).PivotTables("Pivot1").PageFieldOrder = xlOverThenDown
    Application.Worksheets.Item(1).PivotTables("Pivot1").PageFieldWrapCount = 3
}
```

```javascript
/*本示例将数据透视表布局上的顺序设置为：向下处理行，然后向右逐个处理页或页面字段；并弹窗显示此属性。*/
function test() {
    ActiveSheet.PivotTables("Pivot1").PageFieldOrder = xlDownThenOver
    if (ActiveSheet.PivotTables("Pivot1").PageFieldOrder == xlDownThenOver) {
        console.log("数据透视表布局上的顺序为：向下处理行，然后向右逐个处理页或页面字段。")
    }
}
```
