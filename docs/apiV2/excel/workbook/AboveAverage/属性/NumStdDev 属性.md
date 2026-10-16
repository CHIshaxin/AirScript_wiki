---
title: "NumStdDev 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AboveAverage/%E5%B1%9E%E6%80%A7/NumStdDev%20%E5%B1%9E%E6%80%A7.html
---
# NumStdDev (属性)

返回或设置 **[AboveAverage](./../../AboveAverage/AboveAverage 对象.html)** 对象的数字标准偏差。可读/写 **Long** 类型。

## 示例

```javascript
/*本示例显示活动工作表上区域 C1:C10 中第一个（AboveAverage）条件格式的数字标准偏差。*/
function test() {
    let aboveAverage = ActiveSheet.Range("C1:C10").FormatConditions.Item(1)
    console.log(aboveAverage.NumStdDev)
}
```

```javascript
/*本示例设置区域 C1:C10 中第一个（AboveAverage）条件格式的数字标准偏差。*/
function test() {
    let aboveAverage = Application.Range("C1:C10").FormatConditions.Item(1)
    aboveAverage.NumStdDev = 1
}
```
