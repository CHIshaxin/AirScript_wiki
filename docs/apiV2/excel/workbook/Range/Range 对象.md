---
title: "Range 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/Range%20%E5%AF%B9%E8%B1%A1.html
---
# Range (对象)

代表某一单元格、某一行、某一列、某一选定区域（该区域可包含一个或若干连续单元格区域），或者某一三维区域。

## 说明

示例部分中说明了以下用于返回 **[Range](./../Range/Range 对象.html)** 对象的属性和方法：

- **Range** 属性
- **Cells** 属性
- **Range** 和 **Cells**
- **Offset** 属性
- **Union** 方法

## 示例

使用 **Range**(*arg*)（其中 *arg* 为区域名称）可返回一个代表单个单元格或单元格区域的 **[Range](./../Range/Range 对象.html)** 对象。下例将单元格 A1 中的值赋给单元格 A5。

```javascript
/*本示例将 A1 单元格的值赋值给 A5 单元格*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A5").Value2 = Application.Worksheets.Item("Sheet1").Range("A1").Value2
}
```

下例通过为区域 A1:H8 中的每个单元格设置公式，用随机数字填充该区域。如果在不带对象识别符（句点左边的对象）的情况下使用 **Range** 属性，该属性会返回活动表上的一个区域。如果活动表不是工作表，则该方法失败。在使用没有显式对象识别符的 **Range** 属性之前，请先使用 **Activate** 方法激活一个工作表。

```javascript
/*本示例通过为区域 A1:H8 中的每个单元格设置公式，用随机数字填充该区域*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    //Range is on the active sheet
    Application.Range("A1:H8").Formula = "=Rand()"
}
```

下例清除区域名为“*Criteria*”的区域中的内容。 注释：如果用文本参数指定区域地址，必须以 A1 样式记号指定该地址（不能用 R1C1 样式记号）。

```javascript
/*本示例清除区域名为“*Criteria*”的区域中的内容*/
function test() {
    Application.Worksheets.Item(1).Range("Criteria").ClearContents()
}
```

使用 **Cells**(*row*, *column* )（其中 *row* 是行号，*column* 是列标）可返回一个单元格。下例将单元格 A1 赋值为 24。

```javascript
/*本示例将活动工作表上第一行第一个单元格的值设置为 24*/
function test() {
    Application.Worksheets.Item(1).Cells.Item(1, 1).Value2 = 24
}
```

下例设置单元格 A2 的公式。

```javascript
/*本示例在第二行第一列单元格设置公式 =Sum(B1:B5)*/
function test() {
    Application.ActiveSheet.Cells.Item(2, 1).Formula = "=Sum(B1:B5)"
}
```

虽然也可用 Range("A1") 返回单元格 A1，但有时用 **Cells** 属性更为方便，因为对行或列使用变量。下例在 Sheet1 上创建行号和列标。注意，当工作表激活以后，使用 **Cells** 属性时不必明确声明工作表（它将返回活动工作表上的单元格）。 注释：虽然可用 Visual Basic 字符串函数转换 A1 样式引用，但使用 Cells(1, 1) 记号更为简便（而且也是更好的编程习惯）。

```javascript
/*本示例使用循环在单元格填入对应的值*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    for (let i = 1; i <= 5; i++) {
        Application.Cells.Item(1, i + 1).Value2 = 1990 + i
    }
    for (let j = 1; j <= 4; j++) {
        Application.Cells.Item(j + 1, 1).Value2 = "Q" + j
    }
}
```

使用 *expression*.**Cells**(*row* , *column*)（其中 *expression* 是返回 **[Range](./../Range/Range 对象.html)** 对象的表达式，*row* 和 *column* 是相对于该区域左上角的偏移量）可返回区域中的一部分。下例设置单元格 C5 中的公式。

```javascript
/*本示例将 C5:C10 单元格区域中第一行第一列的单元格设置公式 =Rand()*/
function test() {
    Application.Worksheets.Item(1).Range("C5:C10").Cells.Item(1, 1).Formula = "=Rand()"
}
```

使用 **Range**(*cell1, cell2*)（其中 *cell1* 和 *cell2* 是指定起始和终止单元格的 **[Range](./../Range/Range 对象.html)** 对象）可返回一个 **[Range](./../Range/Range 对象.html)** 对象。下例设置单元格区域 A1:J10 的边框线条的样式。 注释：请注意每个 **Cells** 属性之前的句点。如果前导的 **With** 语句应用于 **Cells** 属性，那么这些句点就是必需的。本示例中，句点指示单元格处于工作表一上。如果没有句点，**Cells** 属性将返回活动工作表上的单元格。

```javascript
/*本示例设置单元格 A1:J10 的边框线条样式*/
function test() {
    Application.Worksheets.Item(1).Range(Worksheets.Item(1).Cells.Item(1, 1), Application.Worksheets.Item(1).Cells.Item(10, 10)).Borders.LineStyle = xlThick
}
```

使用 **Offset**(*row, column*)（其中 *row* 和 *column* 为行偏移量和列偏移量）可返回相对于另一区域在指定偏移量处的区域。下例选定位于当前选定区域左上角单元格的向下三行且向右一列处的单元格。由于必须选定位于活动工作表上的单元格，因此必须先激活工作表。

```javascript
/*本示例选择从当前选定区域左上角的单元格下移 3 行和右移 1 列所得的单元格。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    //Can't select unless the sheet is active
    Application.Selection.Offset(3, 1).Range("A1").Select()
}
```

使用 **Union**(*range1, range2*, ...) 可返回多块区域，即该区域由两个或多个连续的单元格区域所组成。下例创建由单元格区域 A1:B2 和 C3:D4 组合定义的对象，然后选定该定义区域。

```javascript
/*本示例选中 A1:B2 和 C3:D4 单元格区域*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let r1 = Range("A1:B2")
    let r2 = Range("C3:D4")
    let myMultiAreaRange = Union(r1, r2)
    myMultiAreaRange.Select()
}
```

如果您处理包含多个区域的选定内容，**Areas** 属性是很有用的。它将多区域选定内容拆分为单个的 **[Range](./../Range/Range 对象.html)** 对象，然后将对象返回为一个集合。您可以对返回的集合使用 **Count** 属性，以查找包含多个区域的选定内容，如下例所示。

```javascript
/*本示例判断如果工作表中选中多个区域，则显示 You cannot carry out this command on multi-area selections*/
function test() {
    let NumberOfSelectedAreas = Application.Selection.Areas.Count
    if (NumberOfSelectedAreas > 1) {
        console.log("You cannot carry out this command " + "on multi-area selections")
    }
}
```
