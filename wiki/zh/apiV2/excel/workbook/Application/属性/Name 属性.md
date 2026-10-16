---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bp9sXz0D
---

# Name (属性)

返回一个代表对象名称的 **String** 值。

## 示例

```javascript
/*本示例显示活动工作簿的名称。*/
function test() {
    console.log("活动工作簿的名称为：" + Application.ActiveWorkbook.Name)
}
```

```javascript
/*本示例显示 Name 的数据类型。*/
function test() {
    console.log("Name的数据类型为：" + typeof (Application.Name))
}
```
