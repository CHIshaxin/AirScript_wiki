---
title: "Size 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E5%B1%9E%E6%80%A7/Size%20%E5%B1%9E%E6%80%A7.html
---
# Size (属性)

返回或设置指定底纹的大小。可读/写 **Single** 类型。

## 示例

```javascript
/*本示例显示第一张工作表中第二个形状阴影的大小。*/
function test() {
    let shadow = Application.Worksheets.Item(1).Shapes.Item(2).Shadow
    console.log(shadow.Size)
}
```

```javascript
/*本示例为活动工作表中第三个形状设置阴影类型并设置阴影大小为90。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(3).Shadow
    shadow.Type = msoShadow16
    shadow.Size = 90
}
```
