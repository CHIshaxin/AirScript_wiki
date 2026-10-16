---
title: "UnMerge 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/UnMerge%20%E6%96%B9%E6%B3%95.html
---
# UnMerge (方法)

将合并区域分解为独立的单元格。

## 示例

```javascript
/*本示例将包含单元格 A3 的合并区域分解。*/
function test() {
    if (Application.Range("A3").MergeCells) {
        Application.Range("A3").MergeArea.UnMerge()
    } else {
        console.log("not merged")
    }
}
```

```javascript
/*本示例将 A1:B3 单元格合并，再通过 B1 单元格将该合并区域分解*/
function test() {
    Range("A1:B3").Merge()
    Range("B1").UnMerge()
}
```
