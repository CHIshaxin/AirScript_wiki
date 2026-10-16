---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DJwWTBx6
---

# Orientation (属性)

返回或设置一个 **Variant** 值，它代表文本方向。

## 说明

此属性的值可以设置为 -90 到 90 度的整数值，或设置为 **[XlOrientation](./../../枚举/XlOrientation 枚举.html)** 常量之一。

## 示例

```javascript
/*此示例显示查找条件的文本方向是否为向下排列。*/
function test() {
    console.log(Application.FindFormat.Orientation == xlDownward)
}
```

```javascript
/*此示例设置查找条件的文本方向为向上排列。*/
function test() {
    Application.FindFormat.Orientation = xlUpward
}
```
