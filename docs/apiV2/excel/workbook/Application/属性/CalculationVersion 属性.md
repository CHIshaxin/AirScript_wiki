---
title: "CalculationVersion 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/CalculationVersion%20%E5%B1%9E%E6%80%A7.html
---
# CalculationVersion (属性)

返回一个数字，其最右侧两位表示计算引擎的次要版本号，而左侧的其他位表示 ET 组件的主要版本号。**Long** 型，只读。

## 说明

如果工作簿是用早期版本 ET 保存的，且该工作簿并未进行完全重新计算，则此属性返回 0。

## 示例

```javascript
/*此示例将当前 ET 的版本与上次计算工作簿的 ET 的版本进行比较。如果两个版本号不同，则此示例将 blnFullCalc 变量设置为true。*/
function test() {
    let blnFullCalc = false
    if (Application.CalculationVersion != Workbooks.Item(1).CalculationVersion) {
        blnFullCalc = true
    }
}
```

```javascript
/*此示例演示 CalculationVersion 属性的值。*/
function test() {
    console.log("CalculationVersion属性的值为：" + Application.CalculationVersion)
}
```
