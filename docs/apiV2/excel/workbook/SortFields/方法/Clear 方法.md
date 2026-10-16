---
title: "Clear 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortFields/%E6%96%B9%E6%B3%95/Clear%20%E6%96%B9%E6%B3%95.html
---
# Clear (方法)

清除所有 **[SortField](./../../SortField/SortField 对象.html)** 对象。

## 示例

```javascript
/*本示例清除活动工作表上所有 SortField 对象。*/
function test() {
    Application.ActiveSheet.Sort.SortFields.Clear()
}
```

```javascript
/*本示例清除活动工作簿上每一张工作表上所有 SortField 对象。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        Application.Sheets.Item(i).Sort.SortFields.Clear()
    }
}
```
