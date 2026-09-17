---
title: "Find 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Find%20%E6%96%B9%E6%B3%95.html
---
# Find (方法)

在区域中查找特定信息。

## 说明

如果未发现匹配项，则返回 **null**。**Find** 方法不影响选定区域或当前活动的单元格。

每次使用此方法后，参数 *LookIn*、*LookAt*、*SearchOrder* 和 *MatchByte* 的设置都将被保存。如果下次调用此方法时不指定这些参数的值，就使用保存的值。设置这些参数将更改 **“查找”** 对话框中的设置，如果省略这些参数，更改 **“查找”** 对话框中的设置将更改使用的保存值。要避免出现这一问题，每次使用此方法时请明确设置这些参数。

使用 **FindNext** 和 **FindPrevious** 方法可重复搜索。

当搜索到指定查找区域的末尾时，此方法将绕回到区域的开始位置继续搜索。发生绕回后，要停止搜索，可保存第一个找到的单元格地址，然后测试后面找到的每个单元格地址是否与其相同。

若要对单元格进行模式更为复杂的搜索，请结合使用 for 循环语句和正则表达式、字符串 startsWith 等方法。例如，下列代码在单元格区域 A1:C5 中使用循环搜索字体名称以 “Cour” 开始的单元格。当 ET 找到匹配单元格以后，就将其字体改为 Times New Roman。

```javascript
let r = Range("A1:C5")
for (i = 1; i <= r.Count; i++) {
let c = r.Item(i)
if (c.Font.Name.startsWith("Cour")) {
c.Font.Name = "Times New Roman"
}
}
```

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| What | any | 必填 | 要搜索的数据。可为字符串或任意 ET 数据类型。 |
| After | any | 可选 | 表示搜索过程将从其之后开始进行的单元格。此单元格对应于从用户界面搜索时的活动单元格的位置。请注意：After 必须是区域中的单个单元格。要记住搜索是从该单元格之后开始的；直到此方法绕回到此单元格时，才对其进行搜索。如果不指定该参数，搜索将从区域的左上角的单元格之后开始。 |
| LookIn | any | 可选 | 信息类型。 |
| LookAt | any | 可选 | 可为以下 XlLookAt 常量之一：xlWhole 或 xlPart。 |
| SearchOrder | any | 可选 | 可为以下 XlSearchOrder 常量之一：xlByRows 或 xlByColumns。 |
| SearchDirection | Default<XlSearchDirection, 1> | 可选 | 搜索的方向。 |
| MatchCase | any | 可选 | 如果为 True，则搜索区分大小写。默认值为 False。 |
| MatchByte | any | 可选 | 只在已经选择或安装了双字节语言支持时适用。如果为 True，则双字节字符只与双字节字符匹配。如果为 False，则双字节字符可与其对等的单字节字符匹配。 |
| SearchFormat | any | 可选 | 搜索的格式。 |

## 示例

```javascript
/*本示例在第一个工作表的单元格区域 A1:A500 中查找包含值 2 的所有单元格，并将这些单元格的值更改为 5。*/
function test() {
    let add = Application.Worksheets.Item(1).Range("A1:A500")
    let c = add.Find(2, undefined, xlValues)
    if (c != null) {
        let firstAddress = c.Address()
        do {
            c.Value2 = 5
            c = add.FindNext(c)
        }while (c != null && c.Address() != firstAddress)
    }
}
```

```javascript
/*本示例将 F1 单元格的公式设置为 =A1+B1，将 F1:F15 单元格根据单元格 F1 进行填充，在 A1:F15 中查找第一个公式为 =A5+B5 的单元格，显示它的地址*/
function test() {
    Range("F1").Formula = '=A1+B1'
    Range("F1").AutoFill(Range("F1:F15"))
    let c = Range("A1:F15").Find("=A5+B5", undefined, xlFormulas)
    console.log(c.Address())
}
```
