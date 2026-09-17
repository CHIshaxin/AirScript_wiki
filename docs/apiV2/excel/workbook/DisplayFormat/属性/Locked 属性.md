---
title: "Locked 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/Locked%20%E5%B1%9E%E6%80%A7.html
---
# Locked (属性)

返回一个值，该值指示当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象是否已锁定。只读。

## 说明

如果该区域已锁定，则返回 **True**；如果在工作表受保护时可以修改该区域，则返回 **False**；如果该区域同时包含锁定和未锁定的单元格，则返回 **Null**。

## 返回值

Variant

## 示例

```javascript
/*本示例显示 G5 单元格是否在用户界面显示时被锁定*/
function test() {
    console.log(Range("G5").DisplayFormat.Locked)
}
```

```javascript
/*本示例判断 D5:D10 单元格是否在用户界面显示时被锁定，并通知用户*/
function test() {
    if (Range("D5:D10").DisplayFormat.Locked) {
        console.log("用户界面显示时被锁定")
    } else {
        console.log("用户界面显示时未被锁定")
    }
}
```
