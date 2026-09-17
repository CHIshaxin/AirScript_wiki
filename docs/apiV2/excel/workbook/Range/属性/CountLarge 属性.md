---
title: "CountLarge 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/CountLarge%20%E5%B1%9E%E6%80%A7.html
---
# CountLarge (属性)

返回一个值，该值表示集合中的对象数。只读。

## 示例

```javascript
/*本示例显示 A1:C3 单元格区域内的单元格数量*/
function test() {
    console.log(Range("A1:C3").CountLarge)
}
```

```javascript
/*本示例选中 C2:F4 单元格，并显示选中区域中单元格的数量*/
function test() {
    Range("C2:F4").select()
    console.log(Selection.CountLarge)
}
```
