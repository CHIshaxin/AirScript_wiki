---
title: "Value 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
---
# Value (属性)

返回一个 **Boolean** 值，它指明是否符合所有的有效性验证条件（即，该区域是否包含有效数据）。

## 示例

```javascript
/*本示例显示单元格 B1 是否包含有效数据。*/
function test() {
    console.log(Range("B1").Validation.Value)
}
```

```javascript
/*本示例显示第一张工作表上单元格 A1 是否包含有效数据。*/
function test() {
    let validation = Application.Worksheets.Item(1).Range("A1").Validation
    console.log(validation.Value)
}
```
