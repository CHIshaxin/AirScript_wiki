---
title: "FilterMode 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AutoFilter/%E5%B1%9E%E6%80%A7/FilterMode%20%E5%B1%9E%E6%80%A7.html
---
# FilterMode (属性)

如果工作表的筛选模式为自动筛选，则返回 **True**。只读 **Boolean** 类型。

## 示例

```javascript
/*本示例显示活动工作表的筛选模式是否为自动筛选。*/
function test() {
    console.log(ActiveSheet.AutoFilter.FilterMode)
}
```

```javascript
/*本示判断第一张工作表的筛选模式是否为自动筛选，并通知用户。*/
function test() {
    if (Worksheets.Item(1).AutoFilter.FilterMode) {
        console.log("工作表的筛选模式是自动筛选")
    } else {
        console.log("工作表的筛选模式不是自动筛选")
    }
}
```
