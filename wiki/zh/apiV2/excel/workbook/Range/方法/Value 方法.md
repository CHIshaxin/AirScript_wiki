---
title: "Value (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Value%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DD88llTb
---

# Value (方法)

返回或设置一个 **Variant** 型，它代表指定单元格的值。

## 说明

用 XML 电子表格文件中的内容设置单元格区域时，只使用工作簿中第一张工作表中的值。无法设置或得到以 XML 电子表格格式表示的不相连的单元格区域。

JSAPI里的Value是个方法，可以进行取值，通过第二个参数可以进行赋值。另外给单元格赋值也可以用Range.Value2属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| RangeValueDataType | any | 可选 | 区域值数据类型。可以为 xlRangeValueDataType 常量。 |
| RHS | any | 可选 | 待赋值该单元格的内容。不传该参数则为取值情况。 |

## 示例

```javascript
/*本示例获取 Sheet1 上 A1 单元格的值，赋给变量"rngValue"。*/
function test() {
    let rngValue = Application.Worksheets.Item("Sheet1").Range("A1").Value()
}
```

```javascript
/*本示例将 A1 单元格的值设置为 a1，将 A1 单元格的值赋值给 A2 单元格*/
function test() {
    Range("A1").Value(null, 'a1')
    Range("A2").Value2 = Range("A1").Value(xlRangeValueDefault)
}
```
