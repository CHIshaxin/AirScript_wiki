---
title: "NameLocal (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/%E5%B1%9E%E6%80%A7/NameLocal%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BWeVljwh
---

# NameLocal (属性)

以用户语言返回对象的名称。只读 **String** 类型。

## 说明

如果样式为内置样式，则此属性以当前系统所使用的地区语言返回样式的名称。

## 示例

```javascript
/*本示例将工作表 Sheet1 上第一张列表的表样式的名称（用户语言表示）赋值给A3单元格。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    Range("A3").Value2 = listObj.TableStyle.NameLocal
}
```

```javascript
/*本示例显示活动工作簿上的表格样式中的第一个样式名称（用户语言表示）。*/
function test() {
    let tableStyles = ActiveWorkbook.TableStyles
    console.log(tableStyles.Item(1).NameLocal)
}
```
