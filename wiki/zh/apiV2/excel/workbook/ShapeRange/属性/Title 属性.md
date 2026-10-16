---
title: "Title (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Title%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CVIJ8H-h
---

# Title (属性)

返回或设置与指定的形状区域关联的可选文本的标题。可读/写。

## 说明

若要设置形状区域的可选文本的描述性文本字符串，请使用 **AlternativeText** 属性。

## 返回值

String

## 示例

```javascript
/*本示例先在第一张工作表中创建shpRange对象，然后将对象的标题设置为“My Shape Range Title”。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.Title = "My Shape Range Title"
}
```

```javascript
/*本示例先在活动工作表中创建shpRange对象，然后显示对象的标题。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([3])
    console.log(shpRange.Title)
}
```
