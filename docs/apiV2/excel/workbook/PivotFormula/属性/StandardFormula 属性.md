---
title: "StandardFormula 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFormula/%E5%B1%9E%E6%80%A7/StandardFormula%20%E5%B1%9E%E6%80%A7.html
---
# StandardFormula (属性)

返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。

## 说明

**StandardFormula**属性主要影响具有日期或数字格式的项目名称。该属性提供了一种方法可指定或查询给定计算项的公式。

**StandardFormula** 属性是国际通用的，**Formula** 属性则不是。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中名为“公式2”的公式的内容。*/
function test() {
    let pvtFormula = ActiveSheet.PivotTables(1).PivotFormulas("公式2")
    console.log(pvtFormula.StandardFormula)
}
```

```javascript
/*本示例显示第一张工作表上数据透视表中名为“公式1”的公式的内容。*/
function test() {
    let pvtFormula = Worksheets.Item(1).Range("I1").PivotTable.PivotFormulas("公式1")
    console.log(pvtFormula.StandardFormula)
}
```
