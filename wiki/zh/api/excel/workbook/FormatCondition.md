---
title: "FormatCondition"
source: https://airsheet.wps.cn/docs/api/excel/workbook/FormatCondition.html
updated: 1789700916000
hash: Csn8_3co
---

# FormatCondition

区域内的某个条件格式

FormatCondition 对象的具体属性和方法请参阅下方的列表。

#### 属性列表

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| [AppliesTo](#appliesto) | [Range](Range.md) | 应用格式规则的单元格区域 |
| [Borders](#borders) | [Border](Border.md) | 返回一个 Borders 集合 |
| [Font](#font) | [Font](Font.md) | 返回一个 Font 对象 |
| [Formula1](#formula1) | String | 返回与条件格式或者数据有效性相关联的值或表达式 |
| [Formula2](#formula2) | String | 返回与条件格式或数据有效性验证第二部分相关联的值或表达式 |
| [Interior](#interior) | Interior | 表示指定对象的内部 |
| [NumberFormat](#numberformat) | String | 单元格的数字格式 |
| [Operator](#operator) | [XlFormatConditionOperator](Enum.md#xlformatconditionoperator) | 条件格式的运算符 |
| [Priority](#priority) | Number | 返回或设置条件格式规则的优先级值 |
| [Type](#type) | [XlFormatConditionType](Enum.md#xlformatconditiontype) | 条件格式对象类型 |

#### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| [Modify()](#modify) | undefined | 更改现有条件格式 |
| [ModifyAppliesToRange()](#modifyappliestorange) | undefined | 设置此格式规则所应用于的单元格区域 |
| [SetFirstPriority()](#setfirstpriority) | undefined | 将此条件格式规则的优先级值设置为“1” |
| [SetLastPriority()](#setlastpriority) | undefined | 将此条件格式规则的优先级值增加“1” |

## AppliesTo

返回一个 Range 对象，该对象指应用格式规则的单元格区域

#### 数据类型

[Range](Range.md) - 区域对象

#### 示例

```js
// 计算A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
// 定义公式:A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const formula1 = `=AND(A9=2, B9>${ave})`
// 指定B列添加条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// 获取条件格式应用的区域
const appliesToRange = formatCondition1.AppliesTo
// 打印目标区域的地址
console.log(appliesToRange.Address())
```

## Borders

返回一个 Borders 集合，该集合表示样式或单元格区域的边框 (包括定义为条件格式) 的一部分的区域

#### 数据类型

[Border](Border.md)集合 - 边框对象集合

#### 示例

```js
// 计算A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
// 定义公式:A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const formula1 = `=AND(A9=2, B9>${ave})`
// 指定B列添加条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// 设置条件格式边框颜色
formatCondition1.Borders.Item(xlDiagonalDown).Color = '#FFFF00'
```

## Font

返回一个 Font 对象，该对象表示指定对象的字体

#### 数据类型

[Font](Font.md) - 字体对象

#### 示例

```js
// 计算A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
// 定义公式:A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const formula1 = `=AND(A9=2, B9>${ave})`
// 指定B列添加条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// 获取条件格式字体对象
const font = formatCondition1.Font
// 打印字体颜色
console.log(font.Color)
```

## Formula1

返回与条件格式或者数据有效性相关联的值或表达式

Formula1 属性可以是常量值、字符串值、单元格引用或公式

#### 数据类型

String - 值或者表达式

#### 示例

```js
// 计算A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
// 定义公式:A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const formula1 = `=AND(A9=2, B9>${ave})`
// 指定B列添加条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// 打印条件格式条件1
console.log(formatCondition1.Formula1)
```

## Formula2

返回与条件格式或数据有效性验证第二部分相关联的值或表达式

仅当数据验证条件格式 Operator 属性为 xlBetween 或 xlNotBetween 时，才使用 Formula2 属性。可为常量值、字符串值、单元格引用或公式

#### 数据类型

String - 值或者表达式

#### 示例

```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(
  xlCellValue,
  xlBetween,
  '90',
  '100'
)
//3.打印条件格式条件二
console.log(betweenCondition.formula2)
```

## Interior

返回一个 Interior 对象，该对象表示指定对象的内部

#### 数据类型

Interior - 内部对象

#### 示例

```js
// 获取B2到G18的range
const range = Range('B2:G18')
// 清除range的条件格式
range.FormatConditions.Delete()
// 通过新增条件格式设置高亮重复项
let formatCondition = range.FormatConditions.AddUniqueValues()
// 将DupeUnique设置为xlDuplicate，代表条件格式显示重复值
formatCondition.DupeUnique = xlDuplicate
const redColor = RGB(255, 0, 0)
// 将重复值的颜色设为红色
formatCondition.Interior.Color = redColor
```

## NumberFormat

在条件格式规则的计算结果为 True 时返回或设置应用于单元格的数字格式

数字格式是使用“单元格格式”对话框的“数字”选项卡上显示的相同格式代码指定的。

您可以使用内置的数字格式，例如 "General" 或者创建自定义数字格式。

#### 数据类型

String - 数字格式

#### 示例

```js
let targetFormulaRange = Range('B:B')
let formatCondition1 = targetFormulaRange.FormatConditions
// 设置满足条件时的格式，B列最大值增加小数点位
formatCondition1.NumberFormat = '0.0'
```

## Operator

返回条件格式的运算符

#### 数据类型

[XlFormatConditionType](Enum.md#xlformatconditionoperator) - 条件格式操作类型枚举

#### 示例

```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '90', '100')
//3.打印条件格式操作符
console.log(betweenCondition.Operator)
```

## Priority

返回或设置条件格式规则的优先级值。 当工作表中存在多个条件格式规则时，优先级确定求值的顺序

#### 数据类型

Number - 优先级

#### 示例

```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '90', '100')
//3.打印条件格式优先级
console.log(betweenCondition.Priority)
```

## Type

返回对象类型

对象类型可以是以下 [XlFormatConditionType](Enum.md#xlformatconditiontype) 枚举中的一个:

- xlCellValue
- xlExpression

#### 数据类型

[XlFormatConditionType](Enum.md#xlformatconditiontype) - 条件格式类型

#### 示例

```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '=90', '=100')
// 打印条件格式类型
console.log(betweenCondition.Type)
```

## Modify()

更改现有条件格式

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Type | [XlFormatConditionType](Enum.md#xlformatconditiontype) | null | 是 | 指定条件格式是基于单元格值还是基于表达式 |
| Operator | [XlFormatConditionOperator](Enum.md#xlformatconditionoperator) | null | 否 | 条件格式运算符，如果 Type 为 xlExpression，则忽略参数 Operator |
| Formula1 | [XlFormatConditionType](Enum.md#xlformatconditiontype) | null | 否 | 与条件格式关联的值或表达式。 可为常量值、字符串值、单元格引用或公式 |
| Formula2 | [XlFormatConditionType](Enum.md#xlformatconditiontype) | null | 否 | 与条件格式关联的值或表达式。 可为常量值、字符串值、单元格引用或公式 |

### 返回值

undefined

#### 示例

```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '=90', '=100')
//3.更改条件格式，改为90到110
betweenCondition.Modify(xlCellValue, xlBetween, '=90', '=110')
```

## ModifyAppliesToRange()

设置此格式规则所应用于的单元格区域

该区域必须采用 A1 引用样式，并且完全包含在作为集合父级的工作表中 FormatConditions 。 可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号），也可以使用货币符号，但会被忽略

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Range | [Range](Range.md) | null | 是 | 此格式规则将应用于的区域 |

### 返回值

undefined

### 示例

```js
// 把B列上的条件格式移动到A列上

//获取B列的条件格式
const formatConditions = Range('B:B').FormatConditions
//设置目的区域
const destRange = Range('A:A')
//获取条件格式的个数
const count = formatConditions.Count
for (let i = 1; i <= count; i++) {
  //修改条件格式的range为目的区域
  formatConditions.Item(i).ModifyAppliesToRange(destRange)
}
```

## SetFirstPriority()

将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则

当工作表中有多个条件格式规则时，如果该规则以前未设置为优先级“1”，则此方法将导致工作表上所有其他现有规则的优先级增加一个

注意

条件格式规则的优先级基于工作表级别应用。

### 返回值

undefined

### 示例

```js
/**问题
 将所有A20:A100中值为"张三"的行中，D20:D100日期为本月的单元格填充为红色
 将所有A11:A20中值为"李四"的行中，D11:D20日期为下月的单元格填充绿色
 */

// 定义公式:A列为张三，D列单元格的年为本年，月为本月时
let formula1 =
  '=AND(A20="张三",YEAR(D20)=YEAR(TODAY()),MONTH(D20)=MONTH(TODAY()))'
// 直接指定D2:D10添加条件格式
let targetFormulaRange = Range('D20:D100')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// SetFirstPriority设置最高优先级，如果设置最低优先级:SetLastPriority
formatCondition1.SetFirstPriority()
// 将满足条件的格式单元格填充红色
formatCondition1.Interior.Pattern = xlPatternSolid
formatCondition1.Interior.Color = RGB(255, 0, 0)

// 定义公式:A列为李四，D列单元格的年为下月所在年，月为下月时
let formula2 =
  '=AND(A11="李四",YEAR(D11)=YEAR(EOMONTH(TODAY(), 1)),MONTH(D11)=MONTH(EOMONTH(TODAY(), 1)))'
// 直接指定D列添加条件格式
const formatCondition2 = Range('D11:D20').FormatConditions.Add(
  xlExpression,
  -1,
  formula2,
  ''
)
// 设置最低优先级
formatCondition2.SetLastPriority()
// 将满足条件的格式单元格填充绿色
formatCondition2.Interior.Pattern = xlPatternSolid
formatCondition2.Interior.Color = RGB(0, 255, 0)
```

## SetLastPriority()

为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则

优先级的实际值将等于工作表上条件格式规则的总数。 如果工作表中有多个条件格式规则，此方法将导致优先级值大于此规则的规则的优先级增加 1

### 返回值

undefined

### 示例

```js
/**问题
 将所有A20:A100中值为"张三"的行中，D20:D100日期为本月的单元格填充为红色
 将所有A11:A20中值为"李四"的行中，D11:D20日期为下月的单元格填充绿色
 */

// 定义公式:A列为张三，D列单元格的年为本年，月为本月时
let formula1 =
  '=AND(A20="张三",YEAR(D20)=YEAR(TODAY()),MONTH(D20)=MONTH(TODAY()))'
// 直接指定D2:D10添加条件格式
let targetFormulaRange = Range('D20:D100')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// SetFirstPriority设置最高优先级，如果设置最低优先级:SetLastPriority
formatCondition1.SetFirstPriority()
// 将满足条件的格式单元格填充红色
formatCondition1.Interior.Pattern = xlPatternSolid
formatCondition1.Interior.Color = RGB(255, 0, 0)

// 定义公式:A列为李四，D列单元格的年为下月所在年，月为下月时
let formula2 =
  '=AND(A11="李四",YEAR(D11)=YEAR(EOMONTH(TODAY(), 1)),MONTH(D11)=MONTH(EOMONTH(TODAY(), 1)))'
// 直接指定D列添加条件格式
const formatCondition2 = Range('D11:D20').FormatConditions.Add(
  xlExpression,
  -1,
  formula2,
  ''
)
// 设置最低优先级
formatCondition2.SetLastPriority()
// 将满足条件的格式单元格填充绿色
formatCondition2.Interior.Pattern = xlPatternSolid
formatCondition2.Interior.Color = RGB(0, 255, 0)
```
