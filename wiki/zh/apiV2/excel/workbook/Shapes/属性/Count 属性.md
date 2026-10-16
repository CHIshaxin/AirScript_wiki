---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CtYi87sd
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例将第一张工作表中所有形状的宽度扩大至原来的1.2倍。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        shapes.Item(i).Width *= 1.2
    }
}
```

```javascript
/*本示例显示活动工作表中形状的数量。*/
function test() {
    let shapes = ActiveSheet.Shapes
    console.log(shapes.Count)
}
```
