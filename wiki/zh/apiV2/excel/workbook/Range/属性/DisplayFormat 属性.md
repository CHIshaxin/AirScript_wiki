---
title: "DisplayFormat (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/DisplayFormat%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CdZ4flIt
---

# DisplayFormat (属性)

返回代表指定区域的显示设置的 **[DisplayFormat](./../../DisplayFormat/DisplayFormat 对象.html)** 对象。只读

## 返回值

DisplayFormat

## 示例

```javascript
/*本示例将 A1 单元格的字体加粗，并使用 DisplayFormat 检查*/
function test() {
    Range("A1").Font.Bold = true
    console.log(Range("A1").DisplayFormat.Font.Bold)
}
```

```javascript
/*本示例将 C5 单元格的内部颜色设置为红色，并使用 DisplayFormat 检查*/
function test() {
    Range("C5").Interior.Color = RGB(255, 0, 0)
    console.log(Range("C5").DisplayFormat.Interior.Color)
}
```
