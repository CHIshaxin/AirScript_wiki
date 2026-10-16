---
title: "Evaluate 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/Evaluate%20%E6%96%B9%E6%B3%95.html
---
# Evaluate (方法)

将一个 ET 名称转换为一个对象或者一个值。

## 说明

该方法可使用下列 ET 名称类型：

- A1 格式引用。可以通过 A1 格式表示法引用单个单元格。所有引用均视为绝对引用。
- 区域。在引用中可以使用区域、交集和联合运算符（分别为冒号、空格和逗号）。
- 定义的名称。可用宏语言指定任何名称。
- 外部引用。可以使用 ! 运算符引用另一工作簿中的单元格或已定义的名称，例如，Evaluate("[BOOK1.XLS]Sheet1!A1")。
- 图表对象。可以指定任何图表对象名称（如“Legend”、“Plot Area”或“Series 1”），以访问该对象的属性和方法。例如，Charts("Chart1").Evaluate("Legend").Font.Name 返回图例中所用字体的名称。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Name | any | 必填 | 公式或使用 ET 命名约定的对象名称。 |

## 示例

```javascript
/*本示例演示用字符串参数调用 Evaluate 方法。*/
function test() {
    Application.Evaluate("A1").Value2 = 25
    let trigVariable = Application.Evaluate("SIN(45)")
    let firstCellInSheet = Application.Workbooks.Item(1).Sheets.Item(1).Evaluate("A1")
    console.log(trigVariable)
    console.log(firstCellInSheet)
}
```

```javascript
/*此示例将Sheet1上 A1 单元格的字体设置为加粗。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let boldCell = "A1"
    Application.Evaluate(boldCell).Font.Bold = true
}
```
