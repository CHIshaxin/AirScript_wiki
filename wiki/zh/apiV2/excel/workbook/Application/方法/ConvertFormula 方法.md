---
title: "ConvertFormula (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/ConvertFormula%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Crpzj-bj
---

# ConvertFormula (方法)

在 A1 和 R1C1 引用样式之间转换、在相对和绝对引用之间转换公式中的单元格引用，或者同时进行这两种转换。**Variant** 类型。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Formula | any | 必填 | 包含要转换的公式的字符串。必须为有效的公式，并且必须以等号开头。 |
| FromReferenceStyle | XlReferenceStyle | 必填 | 公式的引用样式。 |
| ToReferenceStyle | any | 可选 | 指定要返回的引用样式的 **[XlReferenceStyle](./../../枚举/XlReferenceStyle 枚举.html)** 常量。如果省略此参数，则引用样式不变；公式样式保持 *FromReferenceStyle* 参数指定的样式。 |
| ToAbsolute | any | 可选 | 指定转换引用类型的 **[XlReferenceStyle](./../../枚举/XlReferenceStyle 枚举.html)** 常量。如果省略该参数，则引用类型不变。 |
| RelativeTo | any | 可选 | 包含一个单元格的 **[Range](./../../Range/Range 对象.html)** 对象。相对引用与此单元格相对。 |

## 返回值

Variant

## 示例

```javascript
/*本示例将包含 R1C1-样式引用的 SUM 公式转换为包含 A1-样式引用的等价公式，然后显示转换结果。*/
function test() {
    let inputFormula = "=SUM(R10C2:R15C2)"
    console.log(Application.ConvertFormula(inputFormula, xlR1C1, xlA1))
}
```

```javascript
/*本示例将包含 A1-样式引用的 AVERAGE 公式转换为包含 R1C1-样式引用的等价公式，并赋值给D1单元格与E1单元格。*/
function test() {
    let inputFormula = "=AVERAGE(B10:B15)"
    Range("D1").Formula = Application.ConvertFormula(inputFormula, xlA1, xlR1C1, xlA1)
    Range("E1").Formula = Application.ConvertFormula(inputFormula, xlA1, xlR1C1, null, Range("B10"))
}
```
