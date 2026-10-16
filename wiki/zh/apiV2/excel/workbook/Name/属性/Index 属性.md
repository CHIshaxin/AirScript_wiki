---
title: "Index (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DY_vZr7D
---

# Index (属性)

返回 **Number** 值，它代表对象在其同类对象所组成的集合内的索引号。

## 示例

```javascript
/*本示例显示活动工作簿中名称为“test”的索引号。*/
function test() {
    console.log(ActiveWorkbook.Names.Item("test").Index)
}
```

```javascript
/*本示例显示 Sheet1 中名称为“test1”的索引号。*/
function test() {
    let name = Worksheets.Item("Sheet1").Names.Item("test1")
    console.log(`test1的索引号为 ${name.Index}`)
}
```
