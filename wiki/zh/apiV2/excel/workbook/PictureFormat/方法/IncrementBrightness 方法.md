---
title: "IncrementBrightness (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PictureFormat/%E6%96%B9%E6%B3%95/IncrementBrightness%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 8SDA93nP
---

# IncrementBrightness (方法)

按指定数值改变图片亮度。可以使用 **Brightness**属性设置图片的绝对亮度。

## 说明

调整图片亮度不能超出 **Brightness**属性的界限。例如，如果 **Brightness** 属性初始设为 0.9 并将 *Increment* 参数指定为 0.3，则结果亮度为 1.0（**Brightness** 属性的上限）而不是 1.2。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Increment | float | 必填 | 指定图片的 Brightness 属性值的改变量。正值表示图片变亮，负值表示图片变暗。 |

## 示例

```javascript
/*本示例在第一张工作表上创建第一个形状的副本，然后移动并变暗该副本。要执行本示例，第一个形状必须是图片或 OLE 对象。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shape = worksheet.Shapes(1).Duplicate()
    shape.PictureFormat.IncrementBrightness(-0.2)
    shape.IncrementLeft(50)
    shape.IncrementTop(50)
}
```

```javascript
/*本示例将活动工作表的第二个形状图片的亮度增加0.2。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.PictureFormat.IncrementBrightness(0.2)
}
```
