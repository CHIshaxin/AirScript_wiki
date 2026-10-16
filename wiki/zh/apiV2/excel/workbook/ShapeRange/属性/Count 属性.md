---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: PHVexDEd
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例在第一张工作表中按形状的名称生成一组形状的 ShapeRange 集合，并输出集合形状数量。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range(["矩形 1", "椭圆 2"])
    console.log(shpRange.Count)
}
```

```javascript
/*本示例在第一张工作表中按形状的索引生成一组形状的 ShapeRange 集合，并输出集合形状数量。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([3, 2, 1])
    console.log(shpRange.Count)
}
```
