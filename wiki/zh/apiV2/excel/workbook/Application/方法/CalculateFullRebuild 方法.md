---
title: "CalculateFullRebuild (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/CalculateFullRebuild%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BNIKsaO0
---

# CalculateFullRebuild (方法)

对于所有打开的工作簿，强制数据的完整计算并重建从属关系。

## 说明

从属关系是取决于其他单元格的公式。例如，公式“=A1”取决于单元格 A1。**CalculateFullRebuild** 方法类似于重新输入所有公式。

## 示例

```javascript
/*本示例将当前 ET 的版本与上次计算该工作簿的 ET 版本进行比较。如果两个版本不同，则对所有打开工作簿中的数据进行一次完整计算并重建从属关系。*/
function test() {
    if (Application.CalculationVersion !=  Workbooks.Item(1).CalculationVersion) {
        Application.CalculateFullRebuild()
    }
}
```

```javascript
/*本示例中，检查 ET 是否已计算完成。如果已完成，则对所有打开工作簿中的数据进行一次完整计算并重建从属关系。*/
function test() {
    if (Application.CalculationState == xlDone) {
        Application.CalculateFullRebuild()
    }
}
```
