---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFormulas/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BxoYHJBw
---

# Add (方法)

新建数据透视表公式。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Formula | string | 必填 | 新的数据透视表公式。 |
| UseStandardFormula | any | 可选 | 标准数据透视表公式。 |

## 返回值

一个 PivotFormula 对象，它代表新的数据透视表公式。

## 示例

```javascript
/*本示例为第一张工作表的第一张数据透视表添加公式。*/
function test() {
    Worksheets.Item(1).PivotTables(1).PivotFormulas.Add("公式1=张三*0.9", true)
}
```

```javascript
/*本示例为活动工作表的数据透视表添加公式，并显示该数据透视表的第二个公式的名称。*/
function test() {
    let pvtFormulas = ActiveSheet.Range("I1").PivotTable.PivotFormulas
    pvtFormulas.Add("示例数据=张三+15", true)
    console.log(pvtFormulas.Item(2).Value)
}
```
