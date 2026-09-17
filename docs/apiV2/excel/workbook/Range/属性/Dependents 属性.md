---
title: "Dependents 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Dependents%20%E5%B1%9E%E6%80%A7.html
---
# Dependents (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示包含一个单元格所有从属单元格的区域。如果有多个从属单元格，这可能有多个选择（多个 **[Range](./../../Range/Range 对象.html)** 对象）。**Range** 类型，只读。

## 说明

**Dependents** 属性仅在活动工作表中有效，不能追踪远程引用。

## 示例

```javascript
/*本示例选定工作表 Sheet1 中单元格 A1 的从属单元格。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Range("A1").Dependents.Select()
}
```

```javascript
/*本示例将 A5 单元格的公式设置为 =C3+C5，A6 单元格的公式设置为 =C2+C3，显示 C3 单元格的从属单元格的地址*/
function test() {
    Range("A5").Formula = "=C3+C5"
    Range("A6").Formula = "=C2+C3"
    console.log(Range("C3").Dependents.Address())
}
```
