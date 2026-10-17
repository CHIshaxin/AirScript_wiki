---
title: "AutoFilterMode (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/AutoFilterMode%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CubBVGFA
---

# AutoFilterMode (属性)

如果当前在工作表上显示有“自动筛选”下拉箭头，则该值为 **True**。本属性与 **FilterMode** 属性互相独立。**Boolean** 类型，可读写。

## 说明

如果当前显示下拉箭头，则此属性返回 **True**。可以将该属性设置为 **False** 以删除箭头，但不能将其设置为 **True**。使用 **AutoFilter** 方法可筛选列表并显示下拉箭头。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中 AutoFilterMode 属性的当前状态。*/
function test() {
    let isOn
    if (Worksheets.Item("Sheet1").AutoFilterMode) {
        isOn = "On"
    } else {
        isOn = "Off"
    }
    console.log("AutoFilterMode is" + isOn)
}
```

```javascript
/*本示例演示AutoFilterMode属性的数据类型，并将AutoFilterMode属性的值设置为false。*/
function test() {
    console.log(typeof Worksheets.Item(1).AutoFilterMode)
    Worksheets.Item(1).AutoFilterMode = false
}
```
