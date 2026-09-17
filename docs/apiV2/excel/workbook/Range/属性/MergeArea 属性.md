---
title: "MergeArea 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/MergeArea%20%E5%B1%9E%E6%80%A7.html
---
# MergeArea (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象代表包含指定单元格的合并区域。如果指定的单元格不在合并区域内，则该属性返回指定的单元格。只读。**Variant** 类型。

## 说明

**MergeArea** 属性只应用于单个单元格区域。

## 示例

```javascript
/*本示例为包含单元格 A3 的合并区域赋值。*/
function test() {
    let ma = Application.Range("A3").MergeArea
    if (ma.Address() == "$A$3") {
        console.log("not merged")
    } else {
        ma.Cells.Item(1, 1).Value2 = "42"
    }
}
```

```javascript
/*本示例将 A1：A3 单元格合并，并显示包含 A3 单元格的合并区域的地址*/
function test() {
    Range("A1:A3").Merge()
    console.log(Range("A3").MergeArea.Address())
}
```
