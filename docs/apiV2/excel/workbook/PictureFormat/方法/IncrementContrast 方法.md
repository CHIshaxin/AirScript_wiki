---
title: "IncrementContrast 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PictureFormat/%E6%96%B9%E6%B3%95/IncrementContrast%20%E6%96%B9%E6%B3%95.html
---
# IncrementContrast (方法)

按指定值改变图片的对比度。可以使用 **Contrast**属性设置图片的绝对对比度。

## 说明

调整图片的对比度不能超出 **Contrast** 属性的上下限。例如，如果 **Contrast** 属性初始设为 0.9 并将 *Increment* 参数指定为 0.3，则结果对比度为 1.0（**Contrast** 属性的上限）而不是 1.2。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Increment | float | 必填 | 指定图片的 Contrast 属性值的改变量。正值表示对比度增大，负值表示对比度减小。 |

## 示例

```javascript
/*本示例增大第一张工作表上所有对比度未到最大值的图片的对比度。*/
function test() {
    let shapes = Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        if (shapes.Item(i).Type == msoPicture || shapes.Item(i).Type == msoLinkedPicture) {
            shapes.Item(i).PictureFormat.IncrementContrast(0.1)
        }
    }
}
```

```javascript
/*本示例将活动工作表的第二个形状图片的对比度增加0.3。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.PictureFormat.IncrementContrast(0.3)
}
```
