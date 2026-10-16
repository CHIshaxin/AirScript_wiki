---
title: "SetRange 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sort/%E6%96%B9%E6%B3%95/SetRange%20%E6%96%B9%E6%B3%95.html
---
# SetRange (方法)

设置 **[Sort](./../../Sort/Sort 对象.html)** 对象的起始字符和结束字符的位置。

## 说明

仅当对工作表区域应用排序时才能使用 **SetRange**，当区域在表内时不能使用它。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Rng | Range | 必填 | 指定 Sort 集合的区域。 |

## 示例

```javascript
/*本示例设置活动工作表排序区域为C1:C5。*/
function test() {
    Application.ActiveSheet.Sort.SetRange(Range("C1:C5"))
}
```

```javascript
/*本示例新建一张工作表，并为新工作表设置排序区域为A1:A11。*/
function test() {
    let sheet = Application.Sheets.Add()
    sheet.Sort.SetRange(Range("A1:A11"))
}
```
