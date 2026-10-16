---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFormulas/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

包含在集合中的一个 PivotFormula 对象。

## 示例

```javascript
/*本示例显示第一张工作表中第一个数据透视表的第一个公式。*/
function test() {
    console.log(Worksheets.Item(1).PivotTables(1).PivotFormulas.Item(1).Formula)
}
```

```javascript
/*本示例删除第一张工作表上数据透视表中名为“公式2”的公式。*/
function test() {
    let pvtFormula = Worksheets.Item(1).Range("I1").PivotTable.PivotFormulas.Item("公式2")
    pvtFormula.Delete()
}
```
