---
title: "ShowTableStyleColumnStripes 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ShowTableStyleColumnStripes%20%E5%B1%9E%E6%80%A7.html
---
# ShowTableStyleColumnStripes (属性)

**ShowTableStyleColumnStripes** 属性显示镶边列，这些列中的偶数列和奇数列的格式互不相同。这样使数据透视表更易于阅读。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例显示数据透视表的镶边列。*/
function test() {
    ActiveSheet.PivotTables(1).ShowTableStyleColumnStripes = true
}
```

```javascript
/*本示例显示数据透视表 ShowTableStyleColumnStripes 属性的数据类型，并设置为 false。*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).ShowTableStyleColumnStripes)
    ActiveSheet.PivotTables(1).ShowTableStyleColumnStripes = false
}
```
