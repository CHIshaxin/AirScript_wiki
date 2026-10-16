---
title: "CheckAbort (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/CheckAbort%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CtjK5G3I
---

# CheckAbort (方法)

停止 ET 应用程序中的重新计算。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| KeepAbort | any | 可选 | 允许对区域执行重新计算。 |

## 示例

```javascript
/*本示例中，ET 将停止应用程序中的重新计算（不包括单元格 A10）。为查看本示例的结果，假定应用程序中还存在其他计算，这将允许用户查看被指定继续进行重新计算的单元格和其他单元格区域之间的不同。*/
function test() {
    let rngSubtotal = Application.Range("A10")
    //Stop recalculation except for designated cell.
    Application.CheckAbort(rngSubtotal)
}
```

```javascript
/*本示例判断如果有正在执行的计算，ET将停止应用程序中的重新计算（不包括A1:C1区域）。*/
function test() {
    if (Application.CalculationState != xlDone) {
        Application.CheckAbort(Application.Range("A1:C1"))
    }
}
```
