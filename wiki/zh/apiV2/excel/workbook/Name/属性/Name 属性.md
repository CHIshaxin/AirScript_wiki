---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B0wnO0uc
---

# Name (属性)

返回或设置一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示活动工作簿中索引号为2的自定义名称。*/
function test() {
    console.log(ActiveWorkbook.Names.Item(2).Name)
}
```

```javascript
/*本示例显示活动工作簿中第二个名称是否为“test”。*/
function test() {
    console.log(ActiveWorkbook.Names.Item(2).Name == "test")
}
```
