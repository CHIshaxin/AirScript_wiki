---
title: "PivotFormulas 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFormulas/PivotFormulas%20%E5%AF%B9%E8%B1%A1.html
---
# PivotFormulas (对象)

代表数据透视表的公式的集合。每个公式都由一个 **[PivotFormula](./../PivotFormula/PivotFormula 对象.html)** 对象代表。

## 说明

本对象及其相关属性和方法对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效，这是因为它不支持计算字段和计算项。

使用 **PivotFormulas**属性可返回 **[PivotFormulas](./../PivotFormulas/PivotFormulas 对象.html)** 集合。

## 示例

```javascript
/*本示例为活动工作表上的第一张数据透视表创建一个数据透视表公式列表。*/
function test() {
    let r = 10
    for (let i = 1; i <= ActiveSheet.PivotTables(1).PivotFormulas.Count; i++) {
        Cells.Item(r, 1).Value2 = ActiveSheet.PivotTables(1).PivotFormulas.Item(i).Formula
        r++
    }
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的公式的数量。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFormulas.Count)
}
```
