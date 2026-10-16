---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLineCells/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CVkIpAmb
---

# Count (属性)

返回 **[PivotLineCells](./../../PivotLineCells/PivotLineCells 对象.html)** 集合中的项数。只读。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表中行轴上第一条数据透视线的PivotLineCells属性的PivotCell的数量。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotRowAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCells.Count)
}
```
