---
title: "CreateNames 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/CreateNames%20%E6%96%B9%E6%B3%95.html
---
# CreateNames (方法)

在指定区域中依据工作表中的文本标签创建名称。

## 说明

如果未指定 *Top*、*Left*、*Bottom* 或 *Right* 之一，ET 将根据指定区域的形状猜测文本标志。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Top | any | 可选 | 如果该值为 True，则使用顶部行中的标签创建名称。默认值为 False。 |
| Left | any | 可选 | 如果该值为 True，则使用左侧列中的标签创建名称。默认值为 False。 |
| Bottom | any | 可选 | 如果该值为 True，则使用底部行中的标签创建名称。默认值为 False。 |
| Right | any | 可选 | 如果该值为 True，则使用右侧列中的标签创建名称。默认值为 False。 |

## 示例

```javascript
/*本示例给 A1 单元格赋值 a ，并用单元格区域 A1:A3 中的文字创建区域 B1:B3 的名称。注意，指定区域时必须包括那些含有名称文字的单元格，即便只是为单元格区域 B1:B3 创建名称。*/
function test() {
    Range("A1").Value2 = "a"
    let rangeToName = Application.Worksheets.Item("Sheet1").Range("A1:B3")
    rangeToName.CreateNames(null, true)
}
```

```javascript
/*本示例给 C1 单元格赋值 C ，并用 C1 单元格的文本创建 C2 单元格的名称。注意，指定区域时必须包括那些含有名称文字的单元格，即便只是为单元格 C2 创建名称。*/
function test() {
    Range("C1").Value2 = "C"
    let rangeToName = Application.Worksheets.Item("Sheet1").Range("C1:C2")
    rangeToName.CreateNames(true)
}
```
