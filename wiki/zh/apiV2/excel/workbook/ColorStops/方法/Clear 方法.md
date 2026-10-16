---
title: "Clear (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStops/%E6%96%B9%E6%B3%95/Clear%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B_ikQ5Ma
---

# Clear (方法)

清除指定对象。

## 返回值

void

## 示例

```javascript
/*本示例清除选定区域当前 ColorStops。*/
function test() {
    Range("A1:A10").Select()
    let interior = Selection.Interior
    interior.Pattern = xlPatternLinearGradient
    interior.Gradient.Degree = 90
    interior.Gradient.ColorStops.Clear()
}
```

```javascript
/*本示例清除活动单元格当前 ColorStops。*/
function test() {
    ActiveCell.Interior.Gradient.ColorStops.Clear()
}
```
