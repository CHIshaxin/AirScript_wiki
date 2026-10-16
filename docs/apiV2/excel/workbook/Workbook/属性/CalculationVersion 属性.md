---
title: "CalculationVersion 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/CalculationVersion%20%E5%B1%9E%E6%80%A7.html
---
# CalculationVersion (属性)

返回有关工作簿最近一次完全重新计算所使用的 ET 的版本信息。**Long** 型，只读。

## 说明

如果工作簿是用早期版本 ET 保存的，且该工作簿并未进行完全重新计算，则此属性返回 0。

## 示例

```javascript
/*本示例演示有关工作簿最近一次完全重新计算所使用的 ET 的版本信息*/
function test() {
    console.log(ActiveWorkbook.CalculationVersion)
}
```

```javascript
/*本示例演示CalculationVersion属性的数据类型。*/
function test() {
    console.log(typeof ActiveWorkbook.CalculationVersion)
}
```
