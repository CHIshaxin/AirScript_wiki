---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScaleCriteria/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BY7NsJnW
---

# Count (属性)

返回一个 **Long** 值，该值指定色阶条件格式规则的条件的数目。只读。

## 说明

条件的数目等于为色阶分配的阈值的数目。此值要么为 2，要么为 3，具体情况取决于您为条件格式规则使用双色色阶还是使用三色色阶。

## 示例

```javascript
/*本示例显示活动工作表上单元格区域 A1:A10 的第一个色阶条件格式的阈值条件的数量。*/
function test() {
    let criteria = ActiveSheet.Range("A1:A10").FormatConditions.Item(1).ColorScaleCriteria
    console.log(criteria.Count)
}
```
