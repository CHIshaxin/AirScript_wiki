---
title: "Style (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/Style%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DKRhRWDg
---

# Style (属性)

返回一个包含 **[Style](./../../Style/Style 对象.html)** 对象的值，该值表示关联的 **[Range](./../../Range/Range 对象.html)** 对象在当前用户界面中显示时所采用的样式。

## 返回值

Variant

## 示例

```javascript
/*本示例设置 G2 单元格字体的字号为 15 磅，并显示是否设置成功。*/
function test() {
    Range("G2").Style.Font.Size = 15
    console.log(Range("G2").DisplayFormat.Style.Font.Size == 15)
}
```

```javascript
/*本示例显示在当前用户界面中 G2 单元格字体的字号。*/
function test() {
    console.log(Range("G2").DisplayFormat.Style.Font.Size)
}
```
