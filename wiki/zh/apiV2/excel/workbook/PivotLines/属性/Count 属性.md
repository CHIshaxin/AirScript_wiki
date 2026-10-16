---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLines/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BhmGO5Yg
---

# Count (属性)

返回 **[PivotLines](./../../PivotLines/PivotLines 对象.html)** 集合中的项数。只读。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中行轴上数据透视线的数量。*/
function test() {
    let pvtLines = ActiveSheet.Range("I1").PivotTable.PivotRowAxis.PivotLines
    console.log(pvtLines.Count)
}
```
