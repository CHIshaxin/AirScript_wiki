---
title: "Value (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BxLwek4O
---

# Value (属性)

返回或设置一个 **String** 值，它代表规定名称去引用的公式。

## 示例

```javascript
/*本示例显示活动工作簿中索引号为4的名称所引用的公式。*/
function test() {
    console.log(ActiveWorkbook.Names.Item(4).Value)
}
```

```javascript
/*本示例给索引号为3的名称设置所引用的公式。*/
function test() {
    Application.Names.Item(3).Value = "=Sheet1!$A$4"
}
```
