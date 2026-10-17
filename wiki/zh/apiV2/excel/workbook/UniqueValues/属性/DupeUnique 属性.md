---
title: "DupeUnique (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UniqueValues/%E5%B1%9E%E6%80%A7/DupeUnique%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BX6_Y6og
---

# DupeUnique (属性)

返回或设置 **[XlDupeUnique](./../../枚举/XlDupeUnique 枚举.html)** 枚举的常量之一，指定条件格式规则是查找唯一的值还是查找重复的值。

## 示例

```javascript
/*本示例显示活动工作表上区域 C1:C10 中第二个（UniqueValues）条件格式是否查找重复的值。*/
function test() {
    let uniqueValues = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    console.log(uniqueValues.DupeUnique == xlDuplicate)
}
```

```javascript
/*本示例将工作表 Sheet1 上区域 C1:C10 中第二个（UniqueValues）条件格式设置为查找唯一的值。*/
function test() {
    let uniqueValues = Application.Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(2)
    uniqueValues.DupeUnique = xlUnique
}
```
