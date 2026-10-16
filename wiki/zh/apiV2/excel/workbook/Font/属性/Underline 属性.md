---
title: "Underline (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/Underline%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: TzyHBdkn
---

# Underline (属性)

返回或设置应用于字体的下划线类型。可为以下 **XlUnderlineStyle**常量之一。**Variant** 类型，可读写。

## 说明

| **XlUnderlineStyle** 可为以下 **XlUnderlineStyle** 常量之一。 |
| --- |
| **xlUnderlineStyleNone** |
| **xlUnderlineStyleSingle** |
| **xlUnderlineStyleDouble** |
| **xlUnderlineStyleSingleAccounting** |
| **xlUnderlineStyleDoubleAccounting** |

## 示例

```javascript
/*本示例将 Sheet1 中活动单元格的字体设置为单下划线*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Font.Underline = xlUnderlineStyleSingle
}
```

```javascript
/*本示例显示第一个工作表中单元格A6的字体的下划线类型是否为粗双下划线*/
function test() {
    let font = Application.Worksheets.Item(1).Range("A6").Font
    console.log(font.Underline == xlUnderlineStyleDouble)
}
```

``
