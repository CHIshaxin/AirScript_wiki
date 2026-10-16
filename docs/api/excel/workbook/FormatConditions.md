---
title: "条件格式集合(FormatConditions)"
source: https://airsheet.wps.cn/docs/api/excel/workbook/FormatConditions.html
---
# FormatConditions

FormatConditions 集合对象用于控制 Excel 中的条件格式。

条件格式是一种在工作表中格式化单元格的方法，可以根据单元格的值、公式或其他条件自动应用格式，使数据更易于理解和分析。

#### 属性列表

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Count | Number | 返回 FormatConditions 集合中的对象数 |

#### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Add() | [FormatCondition](FormatCondition.md) | 向 FormatConditions 集合中添加一个条件格式 |
| AddAboveAverage() | AboveAverage | 返回表示指定区域的条件格式规则的新 AboveAverage 对象 |
| AddIconSetCondition() | IconSetCondition | 代表指定区域的图标集条件格式规则 |
| AddColorScale() | ColorScale | 该条件格式规则使用单元格颜色中的渐变来指示所选区域中包含的单元格值的相对差异 |
| AddTop10() | Top10 | 该条件格式可以根据指定的截止值查找单元格区域中的最高值和最低值 |
| AddUniqueValues() | UniqueValues | 返回表示指定区域的条件格式规则的新 UniqueValues 对象 |
| Delete() | undefined | 删除该区域下的条件格式 |
| Item() | [FormatCondition](FormatCondition.md) | 从条件格式集合中返回一个条件格式对象 |

## Count

返回 FormatConditions 集合中的对象数

#### 数据类型

Number - 对象数量

#### 示例

```js
// 获取FormatConditions对象
const formatConditions = Range('A:A').FormatConditions
// 获取条件格式对象数量
const count = formatConditions.Count

console.log(count)
```

## Add()

向 FormatConditions 集合中添加一个条件格式

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Type | [XlFormatConditionType](Enum.md#xlformatconditiontype) | null | 是 | 指定条件格式的类型 |
| Operator | [XlFormatConditionOperator](Enum.md#xlformatconditionoperator) | null | 是 | 指定条件格式的运算符，请注意：如果 Type 为 `xlExpression` ，则忽略 Operator 参数 |
| Formula1 | String | null | 是 | 与条件格式关联的值或表达式。 可为常量值、字符串值、单元格引用或公式 |
| Formula2 | String | null | 否 | 当 Operator 为 `xlBetween` 或 `xlNotBetween` 时，表示与条件格式的第二部分关联的值或表达式 (否则，将忽略此参数)。可以是常量值、字符串值、单元格引用或公式 |

#### 返回值

[FormatCondition](FormatCondition.md) - 返回一个 [FormatCondition](FormatCondition.md) 对象，表示添加的条件格式

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
// SetFirstPriority设置最高优先级，如果设置最低优先级:SetLastPriority
formatCondition1.SetFirstPriority()
// 设置满足条件时的格式，设置删除线
formatCondition1.Font.Strikethrough = true
```

## AddAboveAverage()

返回表示指定区域的条件格式规则的新 AboveAverage 对象

对象 AboveAverage 用于在单元格区域中查找高于或低于平均值或标准偏差的值。例如，可在年度业绩评估中查找高于平均业绩的人员。

#### 返回值

AboveAverage 对象

#### 示例

```js
//1.获取D列的条件格式
const formatconditions = Range('D:D').FormatConditions
//2.增加高于平均值条件格式
const aboveAverageCondition = formatconditions.AddAboveAverage()
//3.如果判断高于平均值使用xlAboveAverage，如果判断低于平均值使用xlBelowAverage
aboveAverageCondition.AboveBelow = xlAboveAverage
aboveAverageCondition.SetFirstPriority()
//4.填充颜色设置为黄色
aboveAverageCondition.Font.Color = RGB(255, 255, 0)
```

## AddIconSetCondition()

返回一个新的 IconSetCondition 对象，该对象代表指定区域的图标集条件格式规则。

使用图标集为数据添加注释并将数据分为按阈值隔开的三到五类数据，每种图标均代表某一值范围。

#### 返回值

IconSetCondition 对象

#### 示例

```js
// 用条件格式的图标集将D列和E列标记出来，分值梯度为>=130，90~130，<=90

//应用区域忽略表头
const bIgnoreHeader = true
//获取目标工作簿
const targetWorkbook = ActiveWorkbook
//获取目标工作表
const targetWorksheet = ActiveSheet

const targetColumns = ['D', 'F']
for (let i = 0; i < targetColumns.length; ++i) {
  //选中目标区域
  const targetColumn = targetColumns[i]
  const beginRow = targetWorksheet.UsedRange.Row + (bIgnoreHeader ? 1 : 0)
  const endRow =
    targetWorksheet.UsedRange.Row + targetWorksheet.UsedRange.Rows.Count - 1
  const targetRangeStr = targetColumn + beginRow + ':' + targetColumn + endRow
  let targetRange = targetWorksheet.Range(targetRangeStr)
  //选中条件格式应用区域
  targetRange.Select()
  //在targetRange上添加新的IconSetCondition对象，代表图标集条件格式
  let iconSetCondition = targetRange.FormatConditions.AddIconSetCondition()
  //将此条件格式规则的优先级值设置为 1
  iconSetCondition.SetFirstPriority()
  //将三个交通灯图标应用于条件格式
  iconSetCondition.IconSet = targetWorkbook.IconSets.Item(xl3TrafficLights1)

  const maxIndex = 3
  const middleIndex = 2
  //设置最大的条件为单元格值>=130
  let IconCriterion3 = iconSetCondition.IconCriteria.Item(maxIndex)
  IconCriterion3.Type = xlConditionValueNumber
  IconCriterion3.Operator = xlGreaterEqual
  IconCriterion3.Value = 130

  //设置中间条件为单元格值>90
  let IconCriterion2 = iconSetCondition.IconCriteria.Item(middleIndex)
  IconCriterion2.Type = xlConditionValueNumber
  IconCriterion2.Operator = xlGreater
  IconCriterion2.Value = 90
}
```

## AddColorScale()

返回一个新的 ColorScale 对象，该对象表示条件格式规则，该规则使用单元格颜色中的渐变来指示所选区域中包含的单元格值的相对差异。

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| ColorScaleType | Number | null | 是 | 色阶的类型，例如传 2 就代表双色刻度，传 3 就代表三色刻度 |

#### 返回值

ColorScale 对象

#### 示例

```js
// 对B列按照最低值设置双色刻度，其中最小刻度为红色，最大刻度为绿色。

//获取B列的条件格式
const formatConditions = Range('B:B').FormatConditions
//增加色阶条件格式，参数2代表2阶，3代表3阶
const color2ScaleCondition = formatConditions.AddColorScale(2)
color2ScaleCondition.SetFirstPriority()
//获取2阶色阶的第一个色阶
const colorScaleCriteria1 = color2ScaleCondition.ColorScaleCriteria.Item(1)
//设置类型。xlConditionValueLowestValue代表最低值。xlConditionValueHighestValue代表最高值。xlConditionValuePercent代表使用百分之比。xlConditionValuePercentile代表使用百分点值。xlConditionValueNumber代表使用数字
colorScaleCriteria1.Type = xlConditionValueLowestValue
//设置色阶颜色值， 255代表红色
colorScaleCriteria1.FormatColor.Color = RGB(255, 0, 0)
//获取2阶色阶的第二个色阶
const colorScaleCriteria2 = color2ScaleCondition.ColorScaleCriteria.Item(2)
//设置类型。xlConditionValueLowestValue代表最低值。xlConditionValueHighestValue代表最高值。xlConditionValuePercent代表使用百分之比。xlConditionValuePercentile代表使用百分点值。xlConditionValueNumber代表使用数字
colorScaleCriteria2.Type = xlConditionValueHighestValue
//设置颜色为绿色
colorScaleCriteria2.FormatColor.Color = RGB(0, 255, 0)
```

## AddTop10()

返回一个 Top10 对象，该对象表示指定区域的条件格式规则

Top10 使用 对象，可以根据指定的截止值查找单元格区域中的最高值和最低值。 例如，可以查找区域报告中位居前五位的销售产品、客户调查中位居最后百分之十五的产品，或者部门人员分析中位居前 25 位的薪金。

#### 返回值

Top10 对象

#### 示例

```js
// 把L列数值最低的5个标记为绿色字体

//应用区域忽略表头
const bIgnoreHeader = true

//获取目标工作表
const targetWorksheet = ActiveSheet

//选中目标区域
const targetColumn = 'L'
const beginRow = targetWorksheet.UsedRange.Row + (bIgnoreHeader ? 1 : 0)
const endRow =
  targetWorksheet.UsedRange.Row + targetWorksheet.UsedRange.Rows.Count - 1
const targetRangeStr = targetColumn + beginRow + ':' + targetColumn + endRow
let targetRange = targetWorksheet.Range(targetRangeStr)
//选中条件格式应用区域
targetRange.Select()

//基于前十项规则添加条件格式
let top10 = targetRange.FormatConditions.AddTop10()
//标记前5时top10.Percent=false，标记前5%时top10.Percent=true
top10.Percent = false
//设置排名值为5
top10.Rank = 5
//标记前5时top10.TopBottom=xlTop10Top
//标记后5时top10.TopBottom=xlTop10Bottom
top10.TopBottom = xlTop10Bottom
//将条件格式字体标绿
const greenColor = RGB(0, 255, 0)
top10.Font.Color = greenColor
top10.Font.TintAndShade = 0
top10.StopIfTrue = false
```

## AddUniqueValues()

返回表示指定区域的条件格式规则的新 UniqueValues 对象

可以使用 UniqueValues 对象快速可视化包含唯一值或重复值的单元格

#### 返回值

UniqueValues 对象

#### 示例

```js
// 将B2到G18设置高亮重复项背景颜色为红色

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

## Delete()

删除该区域下的条件格式

#### 返回值

undefined

```js
// 删除条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Delete()
```

## Item()

从条件格式集合中返回一个条件格式对象

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Index | Number / String | null | 是 | 目标对象在集合内的索引值，从 1 开始 |

#### 返回值

[FormatCondition](FormatCondition.md) 对象

#### 示例

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
