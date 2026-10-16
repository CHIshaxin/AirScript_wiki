---
title: "FontSize (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/FontSize%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CQijP-Or
---

# FontSize (属性)

以磅为单位返回或设置指定艺术字的字体大小。**Single** 类型，可读写。

## 示例

```javascript
/*本示例将第一张工作表中第四个艺术字的字号设置为 16 磅。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    worksheet.Shapes.Item(4).TextEffect.FontSize = 16
}
```

```javascript
/*此示例显示活动工作表的第二个艺术字的字号。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.TextEffect.FontSize)
}
```
