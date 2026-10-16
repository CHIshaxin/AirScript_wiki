---
title: "Parent (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/%E5%B1%9E%E6%80%A7/Parent%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DFj-LXzb
---

# Parent (属性)

返回指定对象的父对象。只读。

## 示例

```javascript
/*本示例显示活动工作表中数据透视表的第一个字段的第一个数据项的父对象的名称。*/
function test() {
    let pvtItem = ActiveSheet.Range("I1").PivotTable.PivotFields(1).PivotItems(1)
    console.log(pvtItem.Parent.Name)
}
```
