---
title: "AddressLocal (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/AddressLocal%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CW8piRb0
---

# AddressLocal (方法)

以用户语言返回对指定区域的区域引用。**String** 类型，只读。

## 说明

如果引用中包含多个单元格，则 *RowAbsolute* 和 *ColumnAbsolute* 将分别应用于所有的行和所有的列。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| RowAbsolute | any | 可选 | 如果为 **True**，则以绝对引用的形式返回引用的行部分。默认值为 **True**。 |
| ColumnAbsolute | any | 可选 | 如果为 **True**，则以绝对引用的形式返回引用的列部分。默认值为 **True**。 |
| ReferenceStyle | Default | 可选 | 指定 A1-样式或 R1C1-样式引用的 **[XlReferenceStyle](./../../枚举/XlReferenceStyle 枚举.html)** 的常量之一。 |
| External | any | 可选 | 如果为 **True**，则返回外部引用。如果为 **False**，则返回本地引用。默认值为 **False**。 |
| RelativeTo | any | 可选 | 如果 *RowAbsolute* 和 *ColumnAbsolute* 都设置为 **False**，且 *ReferenceStyle* 设置为 **xlR1C1**，则必须包含相对引用的起始点。该参数是一个 **[Range](./../../Range/Range 对象.html)** 对象，它定义了引用的起始点。 |

## 示例

假设使用美国英语语言支持创建了本示例，然后在使用德语语言支持的环境中运行本示例。

```javascript
/*下例对工作表 Sheet1 中的同一单元格地址使用了四种不同的表达方式（以用户语言）。示例中的注释为将要显示于消息框中的地址。*/
function test() {
    var mc;
    mc = Worksheets.Item(1).Cells(1, 1);
    console.log(mc.AddressLocal());// $A$1
    console.log(mc.AddressLocal(false, true));// $A1
    console.log(mc.AddressLocal(undefined, undefined, xlR1C1));// R1C1
    console.log(mc.AddressLocal(false, false, xlR1C1, undefined, Worksheets.Item(1).Cells.Item(3, 3)));// R(-2)C(-2)
}
```

```javascript
/*本示例选中 A3 单元格，显示选中区域的非绝对引用地址（以用户语言）*/
function test() {
    Range("A3").Select()
    console.log(Selection.AddressLocal(false, false))
}
```
