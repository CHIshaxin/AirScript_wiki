---
title: "Weight (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LineFormat/%E5%B1%9E%E6%80%A7/Weight%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CnUbbAeB
---

# Weight (属性)

返回或设置一个 **Single** 值，它代表线条的粗细。

## 示例

```javascript
/*本示例显示活动工作表的第一个形状线条的粗细。*/
function test() {
    let line = ActiveSheet.Shapes.Item(1).Line
    console.log(line.Weight)
}
```

```javascript
/*本示例设置第一张工作表中第二个形状线条的粗细为3磅。*/
function test() {
    let line = Worksheets.Item(1).Shapes.Item(2).Line
    line.Weight = 3
}
```
