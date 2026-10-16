---
title: "Tracking 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/Tracking%20%E5%B1%9E%E6%80%A7.html
---
# Tracking (属性)

返回或设置指定艺术字中每个字符的水平间隔相对于字符宽度的比例。该比例可为从 0 到 5 之间的值。（为该属性指定的值越大，字符间的空间越大；小于 1 的值会产生字符重叠）。可读写。**Single** 类型。

## 说明

以下表格相对于用户界面上的可用设置给出 **Tracking**属性的值。

| 用户界面设置 | 相应的 Tracking 属性值 |
| --- | --- |
| 很紧 | 0.8 |
| 紧密 | 0.9 |
| 正常 | 1.0 |
| 稀疏 | 1.2 |
| 很松 | 1.5 |

## 示例

```javascript
/*本示例将包含文本 Test 的艺术字添加到第一张工作表，并指定字符间距非常紧密。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let newWordArt = worksheet.Shapes.AddTextEffect(msoTextEffect1, "Test", "Arial Black", 36, false, false, 100, 100)
    newWordArt.TextEffect.Tracking = 0.8
}
```

```javascript
/*此示例显示活动工作表的第一个艺术字中每个字符的水平间距相对于字符宽度的比例。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    console.log(shape.TextEffect.Tracking)
}
```
