---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicers/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C4z6NOEo
---

# Count (属性)

返回集合中对象的数量。

## 返回值

Long

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中切片器的数量。*/
function test() {
    console.log(ActiveSheet.Range("I1").PivotTable.Slicers.Count)
}
```
