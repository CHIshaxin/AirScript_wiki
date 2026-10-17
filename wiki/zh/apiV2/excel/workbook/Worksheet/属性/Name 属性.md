---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CrijuDJF
---

# Name (属性)

返回或设置一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例演示第一个工作表的名称。*/
function test() {
    console.log(Worksheets.Item(1).Name)
}
```

```javascript
/*本示例设置活动工作表的名称。*/
function test() {
    let newName = "请输入名称"
    ActiveSheet.Name = newName
}
```
