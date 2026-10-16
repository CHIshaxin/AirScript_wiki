---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Borders/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表单元格区域或样式的边框之一。

## 说明

| **[XlBordersIndex](./../../枚举/XlBordersIndex 枚举.html)** 可为下列 **[XlBordersIndex](./../../枚举/XlBordersIndex 枚举.html)** 常量之一。 |
| --- |
| **xlDiagonalDown** |
| **xlDiagonalUp** |
| **xlEdgeBottom** |
| **xlEdgeLeft** |
| **xlEdgeRight** |
| **xlEdgeTop** |
| **xlInsideHorizontal** |
| **xlInsideVertical** |

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | XlBordersIndex | 必填 | XlBordersIndex 的常量之一。 |

## 示例

```javascript
/*本示例设置 Sheet1 工作表中单元格区域 A1:G1 的底部边界的颜色。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:G1").Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
}
```

```javascript
/*本示例将活动工作簿的第一个样式对象顶部边界设置为绿色。*/
function test() {
    ActiveWorkbook.Styles.Item(1).Borders.Item(xlEdgeTop).Color = RGB(0, 255, 0)
}
```
