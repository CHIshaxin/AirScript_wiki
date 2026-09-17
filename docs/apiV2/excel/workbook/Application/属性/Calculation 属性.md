---
title: "Calculation 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/Calculation%20%E5%B1%9E%E6%80%A7.html
---
# Calculation (属性)

返回或设置一个 **[XlCalculation](./../../枚举/XlCalculation 枚举.html)** 值，它代表计算模式。

## 说明

对于 OLAP （OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。）数据源，此属性仅能返回或被设置为 **xlNormal**。

## 示例

```javascript
/*本示例设置 ET 在将工作簿保存到磁盘之前对工作簿进行计算。*/
function test() {
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = true
}
```

```javascript
/*本示例判断如果Calculation属性的值为xlCalculationSemiautomatic，则将Calculation属性的值设置为xlCalculationAutomatic。*/
function test() {
    if (Application.Calculation == xlCalculationSemiautomatic) {
        Application.Calculation = xlCalculationAutomatic
    }
}
```
