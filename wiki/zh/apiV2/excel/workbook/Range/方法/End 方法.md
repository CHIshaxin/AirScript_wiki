---
title: "End (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/End%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 8v6Jd78s
---

# End (方法)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象代表包含源区域的区域尾端的单元格。等同于按键 End+ 向上键、End+ 向下键、End+ 向左键或 End+ 向右键。 **[Range](./../../Range/Range 对象.html)** 对象，只读。

## 示例

```javascript
/*本示例选定包含单元格 B4 的区域中 B 列顶端的单元格。*/
function test() {
    Application.Range("B4").End(xlUp).Select()
}
```

```javascript
/*本示例选定包含单元格 B4 的区域中第 4 行尾端的单元格。*/
function test() {
    Application.Range("B4").End(xlToRight).Select()
}
```

```javascript
/*本示例将选定区域从单元格 B4 延伸至第四行最后一个包含数据的单元格。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Range("B4", Range("B4").End(xlToRight)).Select()
}
```
