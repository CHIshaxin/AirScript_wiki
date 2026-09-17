---
title: "ShowTableStyleRowStripes 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ShowTableStyleRowStripes%20%E5%B1%9E%E6%80%A7.html
---
# ShowTableStyleRowStripes (属性)

**ShowTableStyleRowStripes** 属性显示镶边行，这些行中的偶数行和奇数行的格式互不相同。这样使数据透视表更易于阅读。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例显示数据透视表镶边行*/
function test() {
    Worksheets.Item(1).PivotTables(1).ShowTableStyleRowStripes = true
}
```

```javascript
/*本示例显示数据透视表 ShowTableStyleRowStripes 属性的数据类型，并设置为 false*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).ShowTableStyleRowStripes)
    ActiveSheet.PivotTables(1).ShowTableStyleRowStripes = false
}
```
