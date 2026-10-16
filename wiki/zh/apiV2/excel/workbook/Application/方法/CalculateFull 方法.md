---
title: "CalculateFull (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/CalculateFull%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BNZUQG-G
---

# CalculateFull (方法)

强制对所有打开工作簿中的数据进行完整计算。

## 示例

```javascript
/*本示例将当前 ET 的版本与上次计算该工作簿的 ET 版本进行比较。如果两个版本不同，则对所有打开工作簿中的数据进行一次完整计算。*/
function test() {
    if (Application.CalculationVersion != Workbooks.Item(1).CalculationVersion) {
        Application.CalculateFull()
    }
}
```

```javascript
/*本示例判断如果存在会触发计算的更改，且还未执行重新计算，则对所有打开工作簿中的数据进行一次完整计算。*/
function test() {
    if (Application.CalculationState == xlPending) {
        Application.CalculateFull()
    }
}
```
