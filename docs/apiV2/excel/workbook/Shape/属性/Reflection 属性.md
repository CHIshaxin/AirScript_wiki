---
title: "Reflection 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Reflection%20%E5%B1%9E%E6%80%A7.html
---
# Reflection (属性)

返回指定形状的 **ReflectionFormat** 对象，该对象包含形状的映像格式属性。只读。

## 示例

```javascript
/*本示例判断活动工作表的第一个形状是否是无反射类型。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    if (shape.Reflection.Type == msoReflectionTypeNone) {
        console.log("反射效果为无反射")
    }else {
        console.log("反射效果不是无反射")
    }
}
```

```javascript
/*本示例给第一个工作表的形状设置不同类型的反射效果。*/
function test() {
    let shape1 = Application.Worksheets.Item(1).Shapes.Item(1)
    let shape2 = Application.Worksheets.Item(1).Shapes.Item(2)
    shape1.Reflection.Type = msoReflectionType1
    shape2.Reflection.Type = msoReflectionType2
}
```
