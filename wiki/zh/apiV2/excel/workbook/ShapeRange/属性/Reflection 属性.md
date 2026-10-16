---
title: "Reflection (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Reflection%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: _MdBLcve
---

# Reflection (属性)

返回指定形状区域的 **ReflectionFormat** 对象，该对象包含形状区域的反射格式属性。只读。

## 示例

```javascript
/*本示例判断第一张工作表中新建的ShapeRange对象是否是无反射类型，并显示结果。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([2, 1])
    if (shpRange.Reflection.Type == msoReflectionTypeNone) {
        console.log("反射效果为无反射")
    }else {
        console.log("反射效果不是无反射")
    }
}
```

```javascript
/*本示例将活动工作表中新建的shpRange1和shpRange2对象设置不同类型的反射效果。*/
function test() {
    let shpRange1 = ActiveSheet.Shapes.Range([1])
    let shpRange2 = ActiveSheet.Shapes.Range([2, 3])
    shpRange1.Reflection.Type = msoReflectionType1
    shpRange2.Reflection.Type = msoReflectionType3
}
```
