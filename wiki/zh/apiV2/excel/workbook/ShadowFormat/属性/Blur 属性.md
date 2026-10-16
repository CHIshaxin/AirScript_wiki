---
title: "Blur (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E5%B1%9E%E6%80%A7/Blur%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: ZjOo1MN4
---

# Blur (属性)

返回或设置指定底纹的模糊度。可读/写 **Single** 类型。

## 示例

```javascript
/*本示例显示第一张工作表中第二个形状阴影的模糊度。*/
function test() {
    let shadow = Application.Worksheets.Item(1).Shapes.Item(2).Shadow
    console.log(shadow.Blur)
}
```

```javascript
/*本示例将活动工作表中第一个形状阴影的模糊度设置为8。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(1).Shadow
    shadow.Blur = 8
}
```
