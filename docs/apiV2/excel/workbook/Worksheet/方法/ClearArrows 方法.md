---
title: "ClearArrows 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/ClearArrows%20%E6%96%B9%E6%B3%95.html
---
# ClearArrows (方法)

清除指定工作表的追踪箭头。使用审核功能可添加追踪箭头。

## 示例

```javascript
/*本示例清除 Sheet1 的追踪箭头。*/
function test() {
    Worksheets.Item("Sheet1").ClearArrows()
}
```

```javascript
/*本示例清除所有工作表中的追踪箭头。*/
function test() {
    for (let i = 1; i <= Worksheets.Count; i++) {
        Worksheets.Item(i).ClearArrows()
    }
}
```
