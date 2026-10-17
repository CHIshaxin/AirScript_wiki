---
title: "Enum"
source: https://airsheet.wps.cn/docs/api/excel/databook/Enum.html
updated: 1789700916000
hash: NlGse3aY
---

# Enum

枚举类型，存放在 Application 下

## XlAboveBelow

指定值是高于还是低于平均值

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlAboveAverage | 0 | 高于平均值 |
| xlBelowAverage | 1 | 低于平均值 |
| xlEqualAboveAverage | 2 | 等于平均值 |

## XlAutoFillType

根据源区域的内容，指定目标区域的填充方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlFillDefault | 0 | 确定用于填充目标区域的值和格式 |
| xlFillCopy | 1 | 将源区域的值和格式复制到目标区域，如有必要可重复执行 |
| xlFillSeries | 2 | 将源区域中的值扩展到目标区域中，形式为系列（如，“1, 2” 扩展为 “3, 4, 5”）。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillFormats | 3 | 只将源区域的格式复制到目标区域，如有必要可重复执行 |
| xlFillValues | 4 | 只将源区域的值复制到目标区域，如有必要可重复执行 |
| xlFillDays | 5 | 将星期中每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillWeekdays | 6 | 将工作周每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillMonths | 7 | 将月名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillYears | 8 | 将年从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlLinearTrend | 9 | 将数值从源区域扩展到目标区域中，假定数字之间是加法关系（如，“1, 2,” 扩展为 “3, 4, 5”，假定每个数字都是前一个数字加上某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlGrowthTrend | 10 | 将数值从源区域扩展到目标区域中，假定源区域的数字之间是乘法关系（如，“1, 2,” 扩展为 “4, 8, 16”，假定每个数字都是前一个数字乘以某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行 |

## XlBordersIndex

指定要检索的边框

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDiagonalDown | 5 | 从区域中每个单元格的左上角到右下角的边框 |
| xlDiagonalUp | 6 | 从区域中每个单元格的左下角到右上角的边框 |
| xlEdgeLeft | 7 | 区域左边缘的边框 |
| xlEdgeTop | 8 | 区域顶部的边框 |
| xlEdgeBottom | 9 | 区域底部的边框 |
| xlEdgeRight | 10 | 区域右边缘的边框 |
| xlInsideVertical | 11 | 区域中所有单元格的垂直边框（区域以外的边框除外） |
| xlInsideHorizontal | 12 | 区域中所有单元格的水平边框（区域以外的边框除外） |
| xlOutside | 13 | 区域中的 上下左右 |
| xlInside | 14 | 中间区域 |

## XlBorderWeight

指定某一区域周围的边框的粗细

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlMedium | -4138 | 中 |
| xlHairline | 1 | 细线（最细的边框） |
| xlThin | 2 | 细长 |
| xlThick | 4 | 粗（最宽的边框） |

## XlCalcModeType

迭代计算模式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| manual | `manual` | 手动 |
| automatic | `automatic` | 自动 |

## XlChartType

指定图表类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xl3DArea | -4098 | 三维面积图。 |
| xl3DAreaStacked | 78 | 三维堆积面积图。 |
| xl3DAreaStacked100 | 79 | 百分比堆积面积图。 |
| xl3DBarClustered | 60 | 三维簇状条形图。 |
| xl3DBarStacked | 61 | 三维堆积条形图。 |
| xl3DBarStacked100 | 62 | 三维百分比堆积条形图。 |
| xl3DColumn | -4100 | 三维柱形图。 |
| xl3DColumnClustered | 54 | 三维簇状柱形图。 |
| xl3DColumnStacked | 55 | 三维堆积柱形图。 |
| xl3DColumnStacked100 | 56 | 三维百分比堆积柱形图。 |
| xl3DLine | -4101 | 三维折线图。 |
| xl3DPie | -4102 | 三维饼图。 |
| xl3DPieExploded | 70 | 分离型三维饼图。 |
| xlArea | 1 | 面积图 |
| xlAreaStacked | 76 | 堆积面积图。 |
| xlAreaStacked100 | 77 | 百分比堆积面积图。 |
| xlBarClustered | 57 | 簇状条形图。 |
| xlBarOfPie | 71 | 复合条饼图。 |
| xlBarStacked | 58 | 堆积条形图。 |
| xlBarStacked100 | 59 | 百分比堆积条形图。 |
| xlBubble | 15 | 气泡图。 |
| xlBubble3DEffect | 87 | 三维气泡图。 |
| xlColumnClustered | 51 | 簇状柱形图。 |
| xlColumnStacked | 52 | 堆积柱形图。 |
| xlColumnStacked100 | 53 | 百分比堆积柱形图。 |
| xlConeBarClustered | 102 | 簇状条形圆锥图。 |
| xlConeBarStacked | 103 | 堆积条形圆锥图。 |
| xlConeBarStacked100 | 104 | 百分比堆积条形圆锥图。 |
| xlConeCol | 105 | 三维柱形圆锥图。 |
| xlConeColClustered | 99 | 簇状柱形圆锥图。 |
| xlConeColStacked | 100 | 堆积柱形圆锥图。 |
| xlConeColStacked100 | 101 | 百分比堆积柱形圆锥图。 |
| xlCylinderBarClustered | 95 | 簇状条形圆柱图。 |
| xlCylinderBarStacked | 96 | 堆积条形圆柱图。 |
| xlCylinderBarStacked100 | 97 | 百分比堆积条形圆柱图。 |
| xlCylinderCol | 98 | 三维柱形圆柱图。 |
| xlCylinderColClustered | 92 | 簇状柱形圆锥图。 |
| xlCylinderColStacked | 93 | 堆积柱形圆锥图。 |
| xlCylinderColStacked100 | 94 | 百分比堆积柱形圆柱图。 |
| xlDoughnut | -4120 | 圆环图。 |
| xlDoughnutExploded | 80 | 分离型圆环图。 |
| xlLine | 4 | 折线图。 |
| xlLineMarkers | 65 | 数据点折线图。 |
| xlLineMarkersStacked | 66 | 堆积数据点折线图。 |
| xlLineMarkersStacked100 | 67 | 百分比堆积数据点折线图。 |
| xlLineStacked | 63 | 堆积折线图。 |
| xlLineStacked100 | 64 | 百分比堆积折线图。 |
| xlPie | 5 | 饼图。 |
| xlPieExploded | 69 | 分离型饼图。 |
| xlPieOfPie | 68 | 复合饼图。 |
| xlPyramidBarClustered | 109 | 簇状条形棱锥图。 |
| xlPyramidBarStacked | 110 | 堆积条形棱锥图。 |
| xlPyramidBarStacked100 | 111 | 百分比堆积条形棱锥图。 |
| xlPyramidCol | 112 | 三维柱形棱锥图。 |
| xlPyramidColClustered | 106 | 簇状柱形棱锥图。 |
| xlPyramidColStacked | 107 | 堆积柱形棱锥图。 |
| xlPyramidColStacked100 | 108 | 百分比堆积柱形棱锥图。 |
| xlRadar | -4151 | 雷达图。 |
| xlRadarFilled | 82 | 填充雷达图。 |
| xlRadarMarkers | 81 | 数据点雷达图。 |
| xlStockHLC | 88 | 盘高-盘低-收盘图。 |
| xlStockOHLC | 89 | 开盘-盘高-盘低-收盘图。 |
| xlStockVHLC | 90 | 成交量-盘高-盘低-收盘图。 |
| xlStockVOHLC | 91 | 成交量-开盘-盘高-盘低-收盘图。 |
| xlSurface | 83 | 三维曲面图。 |
| xlSurfaceTopView | 85 | 曲面图（俯视图）。 |
| xlSurfaceTopViewWireframe | 86 | 曲面图（俯视线框图）。 |
| xlSurfaceWireframe | 84 | 三维曲面图（线框）。 |
| xlXYScatter | -4169 | 散点图。 |
| xlXYScatterLines | 74 | 折线散点图。 |
| xlXYScatterLinesNoMarkers | 75 | 无数据点折线散点图。 |
| xlXYScatterSmooth | 72 | 平滑线散点图。 |
| xlXYScatterSmoothNoMarkers | 73 | 无数据点平滑线散点图。 |

## XlContainsOperator

指定函数使用的运算符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlContains | 0 | 包含指定的值 |
| xlDoesNotContain | 1 | 不包含指定的值 |
| xlBeginsWith | 2 | 以指定的值开始 |
| xlEndsWith | 3 | 以指定的值结束 |

## XlDeleteShiftDirection

指定如何移动单元格来替换删除的单元格

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlShiftToLeft | -4159 | 单元格向左移动 |
| xlShiftUp | -4162 | 单元格向上移动 |

## XlDirection

指定移动的方向

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDown | -4121 | 下拉 |
| xlToLeft | -4159 | 向左 |
| xlToRight | -4161 | 向右 |
| xlUp | -4162 | 加速 |

## XlDVAlertStyle

指定验证过程中显示的消息框所用的图标

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlValidAlertStop | 1 | 停止图标 |
| xlValidAlertWarning | 2 | 警告图标 |
| xlValidAlertInformation | 3 | 信息图标 |

## XlDVType

指定要对值进行的有效性测试类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlValidateWholeNumber | 1 | 全部数值 |
| xlValidateDecimal | 2 | 数值 |
| xlValidateList | 3 | 值必须存在于指定列表中 |
| xlValidateDate | 4 | 日期值 |
| xlValidateTime | 5 | 时间值 |
| xlValidateTextLength | 6 | 文本长度 |
| xlValidateCustom | 7 | 使用任意公式验证数据有效性 |

## XlExportImgFormatType

导出图片的格式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlImgTypePNG | 0 | 导出 .png |
| xlImgTypeJPG | 1 | 导出 .jpg |
| xlImgTypeBMP | 2 | 导出 .bmp |
| xlImgTypeTIF | 2 | 导出 .tif |

## XlFixedFormatType

指定文件格式的类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlTypePDF | 0 | PDF（.pdf） |
| xlTypeXPS | 1 | XPS（.xps） |
| xlTypeIMG | 2 | IMG（.png、.jpg、.bmp、.tif） |

## XlFormatConditionOperator

指定用于将公式与单元格中的值或 xlBetween 和 xlNotBetween 中的值进行比较，以比较两个公式的运算符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlBetween | 1 | 行间，只在提供了两个公式的情况下才能使用 |
| xlNotBetween | 2 | 不介于。只在提供了两个公式的情况下才能使用 |
| xlEqual | 3 | 平等 |
| xlNotEqual | 4 | 不等于 |
| xlGreater | 5 | 大于 |
| xlLess | 6 | 小于 |
| xlGreaterEqual | 7 | 大于或等于 |
| xlLessEqual | 8 | 小于或等于 |

## XlFormatConditionType

指定条件格式是基于单元格值还是基于表达式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlCellValue | 1 | 单元格值 |
| xlExpression | 2 | 表达式 |
| xlColorScale | 3 | 色阶 |
| xlTop10 | 5 | 前 10 个值 |
| xlUniqueValues | 8 | 唯一值 |
| xlTextString | 9 | 文本字符串 |
| xlBlanksCondition | 10 | 空值条件 |
| xlTimePeriod | 11 | 时间段 |
| xlAboveAverageCondition | 12 | 高于平均值条件 |
| xlNoBlanksCondition | 13 | 无空值条件 |
| xlErrorsCondition | 16 | 错误条件 |
| xlNoErrorsCondition | 17 | 无错误条件 |

## XlHAlign

指定对象的水平对齐方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlHAlignRight | -4152 | 靠右 |
| xlHAlignLeft | -4131 | 靠左 |
| xlHAlignJustify | -4130 | 两端对齐 |
| xlHAlignDistributed | -4117 | 分散对齐 |
| xlHAlignCenter | -4108 | 居中 |
| xlHAlignGeneral | 1 | 按数据类型对齐 |
| xlHAlignFill | 5 | 填充 |
| xlHAlignCenterAcrossSelection | 7 | 跨列居中 |

## XlInsertFormatOrigin

指定从何处复制插入单元格的格式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlFormatFromLeftOrAbove | 0 | 从上方和/或左侧单元格复制格式 |
| xlFormatFromRightOrBelow | 1 | 从下方和/或右侧单元格复制格式 |

## XlInsertShiftDirection

指定插入时单元格的移动方向

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlShiftDown | -4121 | 向下移动单元格 |
| xlShiftToRight | -4161 | 向上移动单元格 |

## XlLineStyle

指定边框的线条样式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlLineStyleNone | -4142 | 无线 |
| xlDouble | -4119 | 双线 |
| xlDot | -4118 | 点式线 |
| xlDash | -4115 | 虚线 |
| xlContinuous | 1 | 实线 |
| xlDashDot | 4 | 点划相间线 |
| xlDashDotDot | 5 | 划线后跟两个点 |
| xlSlantDashDot | 13 | 倾斜的划线 |

## XlPasteSpecialOperation

XlPaste 特殊操作

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlPasteSpecialOperationAdd | 0x1 | 复制的数据与目标单元格中的值相加 |
| xlPasteSpecialOperationDivide | 0x4 | 复制的数据除以目标单元格中的值 |
| xlPasteSpecialOperationMultiply | 0x3 | 复制的数据乘以目标单元格中的值 |
| xlPasteSpecialOperationNone | 0x0 | 粘贴操作中不执行任何计算 |
| xlPasteSpecialOperationSubtract | 0x2 | 复制的数据减去目标单元格中的值 |

## XlPasteType

粘贴类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlPasteFormulas | 0x2 | 公式 |
| xlPasteAllExceptBorders | 0x4 | 无边框 |
| xlPasteColumnWidths | 0x5 | 保留源列宽 |
| xlPasteValues | 0x3 | 值 |
| xlPasteValuesAndNumberFormats | 0x7 | 值和数字格式 |
| xlPasteFormats | 0x8 | 格式 |
| xlPastePasteAll | 0x0 | 粘贴 |
| xlPasteComments | 0x9 | 批注 |
| xlPasteValidation | 0xa | 数据有效性验证 |

## XlReferenceStyle

指定引用样式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlR1C1 | -4150 | 使用 xlR1C1 返回 R1C1 样式的引用 |
| xlA1 | 1 | 默认值。 使用 xlA1 返回 A1 样式的引用 |

## XlRowCol

指定对应于特定数据系列的数值是处于行中还是列中

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlRows | 1 | 数据系列在列中 |
| xlColumns | 2 | 数据系列在行中 |

## XlSheetType

指定工作表类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlWorksheet | -4167 | 工作表 |
| xlDialogSheet | -4116 | 对话框工作表 |
| xlChart | -4109 | 图表 |
| xlExcel4MacroSheet | 3 | Excel 版本 4 宏工作表 |
| xlExcel4IntlMacroSheet | 4 | Excel 版本 4 国际宏工作表 |

## XlTimePeriods

指定时间段

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlToday | 0 | 今天 |
| xlYesterday | 1 | 昨天 |
| xlLast7Days | 2 | 过去 7 天 |
| xlThisWeek | 3 | 本周 |
| xlLastWeek | 4 | 上周 |
| xlLastMonth | 5 | 上月 |
| xlTomorrow | 6 | 明天 |
| xlNextWeek | 7 | 下周 |
| xlNextMonth | 8 | 下月 |
| xlThisMonth | 9 | 本月 |

## XlUnderlineStyle

指定应用于字体的下划线类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlUnderlineStyleDouble | -4119 | 粗双下划线 |
| xlUnderlineStyleDoubleAccounting | 5 | 紧靠在一起的两条细下划线 |
| xlUnderlineStyleNone | -4142 | 无下划线 |
| xlUnderlineStyleSingle | 2 | 单下划线 |
| xlUnderlineStyleSingleAccounting | 4 | 不支持 |

## XlVAlign

指定对象的垂直对齐方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlVAlignTop | -4160 | 向上 |
| xlVAlignJustify | -4130 | 调整使全行排满 |
| xlVAlignDistributed | -4117 | 一起 |
| xlVAlignCenter | -4108 | 居中 |
| xlVAlignBottom | -4107 | 向下 |

## XlXLMMacroType

指定在工作表中，名称引用哪种宏，或名称是否引用宏

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlFunction | 1 | 自定义函数 |
| xlCommand | 2 | 自定义命令 |
| xlNotXLM | 3 | 非宏 |

## XlAutoFilterOperator

当前工作表中，指定筛选的类型，也用于指定用于关联两个筛选条件的操作符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlAnd | 1 | Criteria1 和 Criteria2 的逻辑与 |
| xlOr | 2 | Criteria1 或 Criteria2 的逻辑或 |
| xlTop10Items | 3 | 显示最高值项 (在 Criteria1 中指定的项目数) |
| xlBottom10Items | 4 | 显示最低值项 (在 Criteria1 中指定的项目数) |
| xlTop10Percent | 5 | 显示最高值项 (Criteria1 中指定的百分比) |
| xlBottom10Percent | 6 | 显示最低值项 (在 Criteria1 中指定的百分比) |
| xlFilterValues | 7 | 筛选值 |
| xlFilterCellColor | 8 | 单元格颜色 |
| xlFilterFontColor | 9 | 字体颜色 |
| xlFilterIcon | 10 | 筛选图标 |
| xlFilterDynamic | 11 | 动态筛选 |

## XlYesNoGuess

指定第一行是否包含标题。 对数据透视表进行排序时，不能使用该参数

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlGuess | 0 | 自动判断 Excel 是否有表头 |
| xlYes | 1 | 默认值。 应对整个区域进行排序 |
| xlNo | 2 | 不应对整个区域进行排序 |

## XlSortOrientation

指定排序方向

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlSortColumns | 1 | 按列排序 |
| xlSortRows | 2 | 按行排序，此值为默认值 |

## XlSortMethod

指定排序类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlPinYin | 1 | 按字符的汉语拼音顺序排序，此值为默认值 |
| xlStroke | 2 | 按每个字符的笔划数排序 |

## XlSortDataOption

指定文本的排序方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlSortNormal | 0 | 分别对数字和文本数据进行排序，此值为默认值 |
| xlSortTextAsNumbers | 1 | 将文本作为数字型数据进行排序 |

## XlSortOn

指定数据的排序参数

| 字段 | 值 | 释义 |
| --- | --- | --- |
| SortOnCellColor | 1 | 单元格颜色 |
| SortOnFontColor | 2 | 字体颜色 |
| SortOnIcon | 3 | 图标 |
| SortOnValues | 0 | 值 |

## XlSortOrder

为指定字段或范围指定排序顺序

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlAscending | 1 | 默认值，按升序对指定字段排序 |
| xlDescending | 2 | 按降序对指定字段排序 |

## XlTextParsingType

指定要导入到查询表中的文本文件中的数据的列格式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDelimited | 1 | 默认值，指示文件由分隔符分隔 |
| xlFixedWidth | 2 | 指示将文件中的数据排列在固定宽度的列中 |

## XlTextQualifier

指定用于指定文本的分隔符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlTextQualifierDoubleQuote | 1 | 双引号 (") |
| xlTextQualifierNone | -4142 | 无分隔符 |
| xlTextQualifierSingleQuote | 2 | 单引号 (') |

## XlColumnDataType

指定列的分列方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDMYFormat | 4 | DMY 日期格式 |
| xlDYMFormat | 7 | DYM 日期格式 |
| xlEMDFormat | 10 | EMD 日期格式 |
| xlGeneralFormat | 1 | 常规 |
| xlMDYFormat | 3 | MDY 日期格式 |
| xlMYDFormat | 6 | MYD 日期格式 |
| xlSkipColumn | 9 | 列未分列 |
| xlTextFormat | 2 | 文本 |
| xlYDMFormat | 8 | YDM 日期格式 |
| xlYMDFormat | 5 | YMD 日期格式 |
