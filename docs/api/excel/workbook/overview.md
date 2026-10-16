---
title: "API总览"
source: https://airsheet.wps.cn/docs/api/excel/workbook/overview.html
---
# API总览

### 表格实例(Application)

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [ActiveSheet](Application.md#activesheet) | [Sheet](Sheet.md) | 当前的活动工作表 |
| [Sheets](Application.md#sheets) | [Sheets](Application.md#sheets) | 当前文件的所有工作表 |
| [ActiveWorkbook](Application.md#activeworkbook) | [Workbook](Workbook.md) | 当前的文档 |
| [Selection](Application.md#selection) | [Range](Range.md) | 当前的选区对象 |
| [Cells](Application.md#cells) | [Range](Range.md) | 当前工作表所有单元格 |
| [Columns](Application.md#columns) | [Range](Range.md) | 当前工作表所有列 |
| [Rows](Application.md#rows) | [Range](Range.md) | 当前工作表所有行 |
| [FileInfo](Application.md#fileinfo) | Object | 当前文档的信息 |
| [UserInfo](Application.md#userinfo) | Object | 当前文档的用户信息 |
| [Enum](Application.md#enum) | [Enum](Enum.md) | 所有的枚举类型 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Range(address)](Application.md#range) | [Range](Range.md) | 获取当前 ActiveSheet 的某个区域(address 指定) |
| [Sheets(name)](Application.md#sheets-1) | [Sheet](Sheet.md) | 获取名称为 name 的工作表 |

### 工作簿(Workbook)

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [ActiveSheet](Workbook.md#activesheet) | [Sheet](Sheet.md) | 工作簿中的活动工作表 |
| [Sheets](Workbook.md#sheets) | [Sheets](Application.md#sheets) | 工作表的所有对象集合 |
| [ReadOnly](Workbook.md#readonly) | Boolean | 文档是否只读 |
| [ReadOnlyComment](Workbook.md#readonlycomment) | Boolean | 文档是否只读可评论的权限 |
| [SupportReadOnlyComment](Workbook.md#supportreadonlycomment) | Boolean | 文档是否支持只读可评论权限 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Save()](Workbook.md#save) | String(JSON) | 保存文件的改动 |
| [GetComments()](Workbook.md#getcomments) | String(JSON) | 获取整个 Workbook 的评论 |
| [ExportAsFixedFormat()](Workbook.md#exportasfixedformat) | String(JSON) | 导出整个表格的 PDF 或者 Img 图片 |

### 工作表(Sheet)

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [Id](Sheet.md#id) | String | 该工作表的 Id |
| [Name](Sheet.md#name) | String | 该工作表的名称 |
| [Index](Sheet.md#index) | Number | 该工作表在所有工作表的索引值 |
| [Cells](Sheet.md#cells) | [Range](Range.md) | 该工作表上所有单元格的集合 |
| [Columns](Sheet.md#columns) | [Range](Range.md) | 该工作表上所有列的集合 |
| [Rows](Sheet.md#rows) | [Range](Range.md) | 该工作表上所有行的集合 |
| [UsedRange](Sheet.md#usedrange) | [Range](Range.md) | 该工作表的使用范围 |
| [Visible](Sheet.md#visible) | Boolean | 该工作表是否可见 |
| [Type](Sheet.md#type) | String | 该工作表的类型 |
| [Hyperlinks](Sheet.md#hyperlinks) | [Hyperlinks](Sheet.md#hyperlinks) | 该工作表上所有超链接的集合 |
| [Shapes](Sheet.md#shapes) | [Shapes](Sheet.md#shapes) | 该工作表上所有图形的集合 |
| [Sort](Sheet.md#sort) | [Sort](Sheet.md#/api/excel/workbook/Sort) | 该工作表上排序对象 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Range()](Sheet.md#range) | [Range](Range.md) | 一个单元格或单元格区域 |
| [Cells()](Sheet.md#cells-1) | [Range](Range.md) | 该工作表上的某个单元格 |
| [Activate()](Sheet.md#activate) | undefined | 切换(激活)工作表 |
| [Move()](Sheet.md#move) | undefined | 移动工作表 |
| [Delete()](Sheet.md#delete) | undefined | 删除工作表 |

### 区域(Range)

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [Count](Range.md#count) | Number | 区域中单元格的数量 |
| [Text](Range.md#text) | String | 【只读】读取单元格格式化文本 |
| [Value/Value2](Range.md#value) | any/[][]any | 读写单元格中的值 |
| [FormatConditions](Range.md#formatConditions) | [FormatConditions](FormatConditions.md) | 用于控制 Excel 中的条件格式 |
| [Formula](Range.md#formula) | String | 以 A1 样式表示法表示的对象的隐式交叉的公式 |
| [FormulaArray](Range.md#formulaarray) | String | 返回或设置区域的数组公式 |
| [NumberFormat](Range.md#numberformat) | String | 获取或者设置区域的数字格式 |
| [Hidden](Range.md#hidden) | Boolean | 行或者列的隐藏 |
| [Interior.Color](Range.md#interior-color) | String | 内部颜色的十六进制 RGB |
| [HorizontalAlignment](Range.md#horizontalalignment) | [Enum.XlHAlign](Enum.md#xlhalign) | 设置区域的水平对齐方式 |
| [VerticalAlignment](Range.md#verticalalignment) | [Enum.XlVAlign](Enum.md#xlvalign) | 设置区域的垂直对齐方式 |
| [WrapText](Range.md#wraptext) | Boolean | 获取或者设置区域自动换行 |
| [IndentLevel](Range.md#indentlevel) | Number | 单元格缩进 |
| [MergeArea](Range.md#mergearea) | [Range](Range.md) | 单元格的合并区域 |
| [MergeCells](Range.md#mergecells) | Boolean | 区域内是否存在合并的单元格 |
| [Cells](Range.md#cells) | [Range](Range.md) | 区域中的单元格集合 |
| [Rows](Range.md#rows) | [Range](Range.md) | 区域中的行集合 |
| [Columns](Range.md#columns) | [Range](Range.md) | 区域中的列集合 |
| [EntireRow](Range.md#entirerow) | [Range](Range.md) | 区域所在行的整行 |
| [EntireColumn](Range.md#entirecolumn) | [Range](Range.md) | 区域所在列的整列 |
| [Row](Range.md#row) | Number | 区域中第一行的行号 |
| [RowEnd](Range.md#rowend) | Number | 区域中最后一行的行号 |
| [Column](Range.md#column) | Number | 区域中第一列的列号 |
| [ColumnEnd](Range.md#columnend) | Number | 区域中最后一列的列号 |
| [Borders](Range.md#borders) | [Borders](Range.md#borders) | 边框集合对象 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [BorderAround()](Range.md#borderaround) | undefined | 向区域添加边框，并为新边框设置 Border 对象的 Color、LineStyle 和 Weight 属性 |
| [Each()](Range.md#each) | undefined | 遍历选区所选单元格 |
| [Item()](Range.md#item) | [Range](Range.md) | 表示区域中指定的位置 |
| [Offset()](Range.md#offset) | undefined | 对指定区域进行迁移操作 |
| [Replace()](Range.md#replace) | undefined | 对单元格内文本执行替换操作 |
| [Delete()](Range.md#delete) | undefined | 单元格、行、列的删除 |
| [Insert()](Range.md#insert) | undefined | 单元格、行、列的新增 |
| [InsertImage()](Range.md#insertimage) | undefined | 插入单元格图片 |
| [Merge()](Range.md#merge) | undefined | 合并单元格 |
| [UnMerge()](Range.md#unmerge) | undefined | 取消合并单元格 |
| [Address()](Range.md#address) | String | 获取表示使用宏语言的区域引用的 String 值 |
| [AddComment()](Range.md#addcomment) | undefined | 添加评论 |
| [ClearComments()](Range.md#clearcomments) | undefined | 清除区域的评论 |
| [Clear()](Range.md#clear) | undefined | 清空指定区域数据和样式 |
| [ClearContents()](Range.md#clearcontents) | undefined | 清除区域的内容 |
| [ClearFormats()](Range.md#clearformats) | undefined | 清除区域的样式 |
| [ClearHyperlinks()](Range.md#clearhyperlinks) | undefined | 清除区域的超链接样式 |
| [Contain()](Range.md#contain) | Boolean | 判断区域是否重叠 |
| [Copy()](Range.md#copy) | Boolean | 将当前区域对象复制到剪贴板 |
| [Cut()](Range.md#cut) | Boolean | 将当前区域对象粘贴到目标区域 |
| [PasteSpecial()](Range.md#pastespecial) | undefined | 将剪贴板中的内容粘贴到指定的单元格或范围 |
| [FillLeft()](Range.md#fillleft) | undefined | 对指定区域中的单元格执行从右往左填充 |
| [FillRight()](Range.md#fillright) | undefined | 对指定区域中的单元格执行从左往右填充 |
| [FillDown()](Range.md#filldown) | undefined | 对指定区域中的单元格执行从上往下填充 |
| [FillUp()](Range.md#fillup) | undefined | 对指定区域中的单元格执行从下往上填充 |
| [AutoFill()](Range.md#autofill) | undefined | 对指定区域中的单元格执行自动填充 |
| [AutoFilter()](Range.md#autofilter) | undefined | 对指定区域中的单元格执行自动筛选 |
| [AutoFit()](Range.md#autofit) | undefined | 更改区域中的列宽或行高以达到最佳匹配 |
| [Select()](Range.md#select) | undefined | 选择区域 |
| [TextToColumns()](Range.md#texttocolumns) | undefined | 将包含文本的一列单元格分解为若干列 |

### 筛选(AutoFilter)

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [Filters](AutoFilter.md#filters) | [Filters](AutoFilter.md#filters) | 筛选对象集合 |
| [Range](AutoFilter.md#range) | [Range](Range.md) | 筛选区域 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [ApplyFilter()](AutoFilter.md#applyfilter) | undefined | 应用筛选到当前工作表 |
| [ShowAllData()](AutoFilter.md#showalldata) | undefined | 清除所有筛选条件，显示所有数据 |

### 排序(Sort)

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [Header](Sort.md#header) | [XlYesNoGuess](Enum.md#xlyesnoguess) | 指定第一行是否包含标题信息 |
| [MatchCase](Sort.md#matchcase) | Boolean | 是否区分大小写 |
| [Orientation](Sort.md#orientation) | [XlSortOrientation](Enum.md#xlsortorientation) | 指定排序方向 |
| [Rng](Sort.md#rng) | [Range](Range.md) | 返回要执行排序的值的区域 |
| [SortFields](Sort.md#sortfields) | [SortFields](https://airsheet.wps.cn/docs/api/excel/workbook/SortFields.html) | 该对象代表与 **Sort** 对象关联的排序字段的集合 |
| [SortMethod](Sort.md#sortmethod) | [XlSortMethod](Enum.md#xlsortmethod) | 中文排序方法 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Apply()](Sort.md#apply) | undefined | 根据当前应用的排序状态对区域进行排序 |
| [SetRange()](Sort.md#setrange) | undefined | 设置排序发生的范围 |

### 排序字段(SortField)

代表与 \*\*Sort\*\* 对象关联的排序字段对象

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [CustomOrder](SortField.md#customorder) | Variant | 指定对字段进行排序的自定义次序 |
| [DataOption](SortField.md#dataoption) | [XlSortDataOption](Enum.md#xlsortdataoption) | 指定如何在 SortField 对象中指定的区域中对文本进行排序 |
| [Key](SortField.md#key) | [Range](Range.md) | 指定排序字段，该字段确定要排序的值 |
| [Order](SortField.md#order) | [XlSortOrder](Enum.md#xlsortorder) | 确定关键字所指定的值的排序次序 |
| [Priority](SortField.md#priority) | Number | 指定排序字段的优先级 |
| [SortOn](SortField.md#sorton) | [XlSortOn](Enum.md#xlsorton) | 返回或设置要排序的单元格的属性 |
| [SortOnValue](SortField.md#sortonvalue) | Object | 返回针对指定的 SortField 对象执行排序的值 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Delete()](SortField.md#delete) | undefined | 从 SortFields 集合中删除指定的 SortField 对象 |
| [ModifyKey()](SortField.md#modifykey) | undefined | 修改字段中按其排序的键值 |

### 字体(Font)

单元格内字体的属性，包括加粗，颜色，大小，斜体，删除线和下划线。

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [Bold](Font.md#bold) | Boolean | 字体是否加粗 |
| [Color](Font.md#color) | String | 字体的颜色 |
| [Italic](Font.md#italic) | Boolean | 字体是否斜体 |
| [Name](Font.md#name) | String | 字体的名称 |
| [Size](Font.md#size) | Number | 字体的大小 |
| [Strikethrough](Font.md#strikethrough) | Boolean | 字体是否有删除线 |
| [Underline](Font.md#underline) | [Enum.XlUnderlineStyle](Enum.md#xlunderlinestyle) | 字体是否有下划线 |

### 边框(Border)

边框对象，Borders 集合里的某一边框

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [Color](Border.md#color) | Number | 边框的颜色 |
| [Weight](Border.md#weight) | [Enum.XlBorderWeight](Enum.md#xlborderweight) | 边框的粗细 |
| [LineStyle](Border.md#linestyle) | [Enum.XlLineStyle](Enum.md#xllinestyle) | 边框的线条样式 |

### 图形(Shape)

某个工作表上的单个 Shape 图形对象

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [ID](Shape.md#id) | String | 图形 ID |
| [Name](Shape.md#name) | String | 图形名称 |
| [Title](Shape.md#title) | String | 图形标题 |
| [Chart](Shape.md#chart) | [Chart](Chart.md) | 图表对象 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Delete](Shape.md#delete) | undefined | 删除Shape |

### 图表(Chart)

单个图表对象

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [ChartTitle.Text](Chart.md#charttitle-text) | String | 图表标题 |
| [HasTitle](Chart.md#hastitle) | Boolean | 图表标题是否可见 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [SetSourceData()](Chart.md#setsourcedata) | undefined | 为图表设置源数据区域 |

### 超链接(Hyperlink)

单个超链接对象，有转跳地址和显示文本两个属性，两者可以不相等。

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [Address](Hyperlink.md#address) | String | 超链接转跳的地址 |
| [TextToDisplay](Hyperlink.md#texttodisplay) | String | 超链接显示的文本 |

### 条件格式集合(FormatConditions)

FormatConditions 集合对象用于控制 Excel 中的条件格式。

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [Count](FormatConditions.md#count) | Number | 返回 FormatConditions 集合中的对象数 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Add()](FormatConditions.md#add) | [FormatCondition](FormatCondition.md) | 向 FormatConditions 集合中添加一个条件格式 |
| [AddAboveAverage()](FormatConditions.md#addaboveaverage) | AboveAverage | 返回表示指定区域的条件格式规则的新 AboveAverage 对象 |
| [AddIconSetCondition()](FormatConditions.md#addiconsetcondition) | IconSetCondition | 代表指定区域的图标集条件格式规则 |
| [AddColorScale()](FormatConditions.md#addcolorscale) | ColorScale | 该条件格式规则使用单元格颜色中的渐变来指示所选区域中包含的单元格值的相对差异 |
| [AddTop10()](FormatConditions.md#addtop10) | Top10 | 该条件格式可以根据指定的截止值查找单元格区域中的最高值和最低值 |
| [AddUniqueValues()](FormatConditions.md#adduniquevalues) | UniqueValues | 返回表示指定区域的条件格式规则的新 UniqueValues 对象 |
| [Delete()](FormatConditions.md#delete) | undefined | 删除该区域下的条件格式 |
| [Item()](FormatConditions.md#item) | [FormatCondition](FormatCondition.md) | 从条件格式集合中返回一个条件格式对象 |

### 条件格式(FormatCondition)

区域内的某个条件格式

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [AppliesTo](FormatCondition.md#appliesto) | [Range](Range.md) | 应用格式规则的单元格区域 |
| [Borders](FormatCondition.md#borders) | [Border](Border.md) | 返回一个 Borders 集合 |
| [Font](FormatCondition.md#font) | [Font](Font.md) | 返回一个 Font 对象 |
| [Formula1](FormatCondition.md#formula1) | String | 返回与条件格式或者数据有效性相关联的值或表达式 |
| [Formula2](FormatCondition.md#formula2) | String | 返回与条件格式或数据有效性验证第二部分相关联的值或表达式 |
| [Interior](FormatCondition.md#interior) | Interior | 表示指定对象的内部 |
| [NumberFormat](FormatCondition.md#numberformat) | String | 单元格的数字格式 |
| [Operator](FormatCondition.md#operator) | [XlFormatConditionOperator](Enum.md#xlformatconditionoperator) | 条件格式的运算符 |
| [Priority](FormatCondition.md#priority) | Number | 返回或设置条件格式规则的优先级值 |
| [Type](FormatCondition.md#type) | [XlFormatConditionType](Enum.md#xlformatconditiontype) | 条件格式对象类型 |

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Modify()](FormatCondition.md#modify) | undefined | 更改现有条件格式 |
| [ModifyAppliesToRange()](FormatCondition.md#modifyappliestorange) | undefined | 设置此格式规则所应用于的单元格区域 |
| [SetFirstPriority()](FormatCondition.md#setfirstpriority) | undefined | 将此条件格式规则的优先级值设置为“1” |
| [SetLastPriority()](FormatCondition.md#setlastpriority) | undefined | 将此条件格式规则的优先级值增加“1” |

### 数据有效性规则(Validation)

代表工作表区域的数据有效性规则

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Add()](Validation.md#add) | undefined | 新增数据有效性规则 |
| [Modify()](Validation.md#modify) | undefined | 修改数据有效性规则 |
| [Delete()](Validation.md#delete) | undefined | 删除数据有效性规则 |

### 工作表函数(WorksheetFunction)

工作表函数对象是 Excel 中的一个内置对象，它包含了许多常用的 Excel 函数，例如 Sum、Average、Min、Max 等。使用 WorksheetFunction 对象可以在 VBA 中调用这些 Excel 函数，以实现对工作表数据的处理和分析。

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Average()](WorksheetFunction.md#average) | Number | 用于计算指定区域内数字的平均值 |
| [AverageIf()](WorksheetFunction.md#averageif) | Number | 用于计算指定区域内满足给定条件的所有单元格的平均值 |
| [Small()](WorksheetFunction.md#small) | Number | 返回数据集中第 k 个最小值 |
| [Large()](WorksheetFunction.md#large) | Number | 用于在一个数组或一列数据中返回第 k 个最大值 |
| [Min()](WorksheetFunction.md#min) | Number | 用于在一个数组或一列数据中返回最小值 |
| [Max()](WorksheetFunction.md#max) | Number | 用于在一个数组或一列数据中返回最大值 |
| [Sum()](WorksheetFunction.md#sum) | Number | 对某单元格区域中的所有数字求和 |

### 枚举(Enum)

枚举类型，存放在 Application 下
