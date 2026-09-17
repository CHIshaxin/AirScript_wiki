---
title: "CustomSubtotalFunction 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/CustomSubtotalFunction%20%E5%B1%9E%E6%80%A7.html
---
# CustomSubtotalFunction (属性)

返回 **[PivotCell](./../../PivotCell/PivotCell 对象.html)** 对象的自定义分类汇总函数字段的设置。 **[XlConsolidationFunction](./../../枚举/XlConsolidationFunction 枚举.html)** 类型，只读。

## 说明

| **[XlConsolidationFunction](./../../枚举/XlConsolidationFunction 枚举.html)** 可为以下 **[XlConsolidationFunction](./../../枚举/XlConsolidationFunction 枚举.html)** 常量之一。 |
| --- |
| **xlAverage** |
| **xlCount** |
| **xlCountNums** |
| **xlMax** |
| **xlMin** |
| **xlProduct** |
| **xlStDev** |
| **xlStDevP** |
| **xlSum** |
| **xlUnknown** |
| **xlVar** |
| **xlVarP** |

如果 **[PivotCell](./../../PivotCell/PivotCell 对象.html)** 对象类型不是自定义分类汇总，那么 **CustomSubtotalFunction** 属性将返回一个错误。该属性只应用于非 OLAP 源数据（非 OLAP 源数据：非 OLAP 源数据是数据透视表或数据透视图使用的基本数据，该数据来自 OLAP 数据库之外的源。这些数据源包括关系数据库、 ET 工作表中的清单以及文本文件数据库。）。

## 示例

```javascript
/*本示例确定包含单元格 I8 的数据透视项是否使用计数合并函数的自定义分类汇总函数，并通知用户。本示例假定数据透视表位于活动工作表上。*/
function test() {
    try {
        // Determine if custom subtotal function is a count function.
        if (Application.Range("I8").PivotCell.CustomSubtotalFunction == xlCount) {
            console.log("The custom subtotal function is a Count.")
        } else {
            console.log("The custom subtotal function is not a Count.")
        }
    }
    catch (exception) {
        console.log("The selected cell is not a custom subtotal function.")
    }
}
```

```javascript
/*本示例显示包含活动单元格的数据透视项的自定义分类汇总函数是否为“xlAverage”。*/
function test() {
    console.log(ActiveCell.PivotCell.CustomSubtotalFunction == xlAverage)
}
```
