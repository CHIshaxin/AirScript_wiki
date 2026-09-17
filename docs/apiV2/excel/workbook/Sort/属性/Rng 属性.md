---
title: "Rng 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sort/%E5%B1%9E%E6%80%A7/Rng%20%E5%B1%9E%E6%80%A7.html
---
# Rng (属性)

返回要执行排序的值的区域。只读。

## 示例

```javascript
/*本示例显示活动工作表排序区域 A1:A5 中的值。*/
function test() {
    Application.ActiveSheet.Sort.SetRange(Range("A1:A5"))
    let rng = Application.ActiveSheet.Sort.Rng
    console.log(`${rng.Value2}`)
}
```

```javascript
/*本示例选中第一张工作表的排序区域。*/
function test() {
    Application.Sheets.Item(1).Sort.SetRange(Range("C1:C6"))
    Application.Sheets.Item(1).Sort.Rng.Select()
}
```
