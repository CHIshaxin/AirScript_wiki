---
title: "ActiveWorkbook (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/ActiveWorkbook%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CB7OyDm8
---

# ActiveWorkbook (属性)

返回一个 **[Workbook](./../../Workbook/Workbook 对象.html)** 对象，该对象代表活动窗口（顶部窗口）中的工作簿。只读。如果没有打开的窗口，或者“信息”窗口或“剪贴板”窗口为活动窗口，则返回 **null**。

## 示例

```javascript
/*本示例显示活动工作簿的名称。*/
function test() {
    console.log("The name of the active workbook is " + Application.ActiveWorkbook.Name)
}
```

```javascript
/*本示例将活动工作簿的 Saved 属性改为true即代表活动工作簿已保存。*/
function test() {
    Application.ActiveWorkbook.Saved = true
}
```
