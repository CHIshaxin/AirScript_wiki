---
title: "Orientation 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
---
# Orientation (属性)

返回或设置一个 **[XlOrientation](./../../枚举/XlOrientation 枚举.html)** 值，它代表文本方向。

## 示例

```javascript
/*本示例为A1单元格添加文本，并显示A1单元格文本方向是否为文字水平排列。*/
function test() {
    Range("A1").Value2 = "示例文本1"
    console.log(Range("A1").Style.Orientation == xlHorizontal)
}
```

```javascript
/*本示例将B1单元格文本方向设置为向下排列。*/
function test() {
    Range("B1").Style.Orientation = xlDownward
}
```
