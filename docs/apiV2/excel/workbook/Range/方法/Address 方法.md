---
title: "Address 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Address%20%E6%96%B9%E6%B3%95.html
---
# Address (方法)

返回一个 **String** 值，它代表宏语言的区域引用。

## 说明

如果引用包含多个单元格，*RowAbsolute* 和 *ColumnAbsolute* 将应用于所有的行和列。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| RowAbsolute | any | 可选 | 如果为 True，则以绝对引用返回引用的行部分。默认值为 True。 |
| ColumnAbsolute | any | 可选 | 如果为 True，则以绝对引用返回引用的列部分。默认值为 True。 |
| ReferenceStyle | Default<XlReferenceStyle, 1> | 可选 | 引用样式。默认值为 xlA1。 |
| External | any | 可选 | 如果为 True，则返回外部引用。如果为 False，则返回本地引用。默认值为 False。 |
| RelativeTo | any | 可选 | 如果 RowAbsolute 和 ColumnAbsolute 为 False，并且 ReferenceStyle 为 xlR1C1，则必须包括相对引用的起始点。此参数是定义起始点的 Range 对象。 |

## 示例

```javascript
/*下例对工作表 Sheet1 中的同一单元格地址使用了四种不同的表达方式。示例中的注释为将要显示于消息框中的地址。*/
function test() {
    let mc = Application.Worksheets.Item("Sheet1").Cells.Item(1, 1)
    console.log(mc.Address())                                    // $A$1
    console.log(mc.Address(false))                               // $A1
    console.log(mc.Address(undefined, undefined, xlR1C1))          // R1C1
    console.log(mc.Address(false, false, xlR1C1, undefined, Worksheets.Item(1).Cells.Item(3, 3)))      // R[-2]C[-2]
}
```

```javascript
/*本示例选中 A3 单元格，并显示选中区域的非绝对引用地址*/
function test() {
    Range("A3").Select()
    console.log(Selection.Address(false, false))
}
```
