---
title: "CircleInvalid (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/CircleInvalid%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DfD48yn8
---

# CircleInvalid (方法)

对工作表中的无效数据项进行圈释。

## 示例

```javascript
/*本示例对第一张工作表的无效数据项进行圈释。*/
function test() {
    Worksheets.Item(1).CircleInvalid()
}
```

```javascript
/*本示例对活动工作表的无效数据项进行圈释，并通知用户。*/
function test() {
    ActiveSheet.CircleInvalid()
    console.log("已完成对无效数据的圈释")
}
```
