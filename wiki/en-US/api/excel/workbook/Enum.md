---
title: "Enum"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Enum.html
updated: 1789700916000
hash: CI2e6Uol
---

# Enum

Enumeration types stored under Application

## XlAboveBelow

Whether the value is above or below the average.

| Field | Value | Meaning |
|---|---|---|
| xlAboveAverage | 0 | Above average |
| xlBelowAverage | 1 | Below average |
| xlEqualAboveAverage | 2 | Equal to average |

## XlAutoFillType

Specifies how the destination range is filled from the source range.

| Field | Value | Meaning |
|---|---|---|
| xlFillDefault | 0 | Determine values and formats used to fill the destination |
| xlFillCopy | 1 | Copy values and formats from source to destination, repeating as needed |
| xlFillSeries | 2 | Extend source values as a series (e.g. "1, 2" to "3, 4, 5"). Copy formats from source, repeating as needed |
| xlFillFormats | 3 | Copy formats only from source to destination, repeating as needed |
| xlFillValues | 4 | Copy values only from source to destination, repeating as needed |
| xlFillDays | 5 | Extend weekday names from source to destination; copy formats as needed |
| xlFillWeekdays | 6 | Extend workweek day names from source to destination; copy formats as needed |
| xlFillMonths | 7 | Extend month names from source to destination; copy formats as needed |
| xlFillYears | 8 | Extend years from source to destination; copy formats as needed |
| xlLinearTrend | 9 | Extend numbers assuming addition (e.g. "1, 2" to "3, 4, 5"); copy formats as needed |
| xlGrowthTrend | 10 | Extend numbers assuming multiplication (e.g. "1, 2" to "4, 8, 16"); copy formats as needed |

## XlBordersIndex

Border index to retrieve.

| Field | Value | Meaning |
|---|---|---|
| xlDiagonalDown | 5 | Diagonal from top-left to bottom-right of each cell |
| xlDiagonalUp | 6 | Diagonal from bottom-left to top-right of each cell |
| xlEdgeLeft | 7 | Left edge of the range |
| xlEdgeTop | 8 | Top edge of the range |
| xlEdgeBottom | 9 | Bottom edge of the range |
| xlEdgeRight | 10 | Right edge of the range |
| xlInsideVertical | 11 | Vertical interior borders between cells (excluding outside) |
| xlInsideHorizontal | 12 | Horizontal interior borders between cells (excluding outside) |
| xlOutside | 13 | Outer edges on all sides |
| xlInside | 14 | Interior region |

## XlBorderWeight

Border weight around a range.

| Field | Value | Meaning |
|---|---|---|
| xlMedium | -4138 | Medium |
| xlHairline | 1 | Hairline (thinnest) |
| xlThin | 2 | Thin |
| xlThick | 4 | Thick (widest) |

## XlCalcModeType

Calculation mode.

| Field | Value | Meaning |
|---|---|---|
| manual | `manual` | Manual |
| automatic | `automatic` | Automatic |

## XlChartType

Chart type.

| Field | Value | Meaning |
|---|---|---|
| xl3DArea | -4098 | 3-D area |
| xl3DAreaStacked | 78 | 3-D stacked area |
| xl3DAreaStacked100 | 79 | 100% stacked area |
| xl3DBarClustered | 60 | 3-D clustered bar |
| xl3DBarStacked | 61 | 3-D stacked bar |
| xl3DBarStacked100 | 62 | 3-D 100% stacked bar |
| xl3DColumn | -4100 | 3-D column |
| xl3DColumnClustered | 54 | 3-D clustered column |
| xl3DColumnStacked | 55 | 3-D stacked column |
| xl3DColumnStacked100 | 56 | 3-D 100% stacked column |
| xl3DLine | -4101 | 3-D line |
| xl3DPie | -4102 | 3-D pie |
| xl3DPieExploded | 70 | Exploded 3-D pie |
| xlArea | 1 | Area |
| xlAreaStacked | 76 | Stacked area |
| xlAreaStacked100 | 77 | 100% stacked area |
| xlBarClustered | 57 | Clustered bar |
| xlBarOfPie | 71 | Bar of pie |
| xlBarStacked | 58 | Stacked bar |
| xlBarStacked100 | 59 | 100% stacked bar |
| xlBubble | 15 | Bubble |
| xlBubble3DEffect | 87 | 3-D bubble |
| xlColumnClustered | 51 | Clustered column |
| xlColumnStacked | 52 | Stacked column |
| xlColumnStacked100 | 53 | 100% stacked column |
| xlConeBarClustered | 102 | Clustered cone bar |
| xlConeBarStacked | 103 | Stacked cone bar |
| xlConeBarStacked100 | 104 | 100% stacked cone bar |
| xlConeCol | 105 | 3-D cone column |
| xlConeColClustered | 99 | Clustered cone column |
| xlConeColStacked | 100 | Stacked cone column |
| xlConeColStacked100 | 101 | 100% stacked cone column |
| xlCylinderBarClustered | 95 | Clustered cylinder bar |
| xlCylinderBarStacked | 96 | Stacked cylinder bar |
| xlCylinderBarStacked100 | 97 | 100% stacked cylinder bar |
| xlCylinderCol | 98 | 3-D cylinder column |
| xlCylinderColClustered | 92 | Clustered cylinder column |
| xlCylinderColStacked | 93 | Stacked cylinder column |
| xlCylinderColStacked100 | 94 | 100% stacked cylinder column |
| xlDoughnut | -4120 | Doughnut |
| xlDoughnutExploded | 80 | Exploded doughnut |
| xlLine | 4 | Line |
| xlLineMarkers | 65 | Line with markers |
| xlLineMarkersStacked | 66 | Stacked line with markers |
| xlLineMarkersStacked100 | 67 | 100% stacked line with markers |
| xlLineStacked | 63 | Stacked line |
| xlLineStacked100 | 64 | 100% stacked line |
| xlPie | 5 | Pie |
| xlPieExploded | 69 | Exploded pie |
| xlPieOfPie | 68 | Pie of pie |
| xlPyramidBarClustered | 109 | Clustered pyramid bar |
| xlPyramidBarStacked | 110 | Stacked pyramid bar |
| xlPyramidBarStacked100 | 111 | 100% stacked pyramid bar |
| xlPyramidCol | 112 | 3-D pyramid column |
| xlPyramidColClustered | 106 | Clustered pyramid column |
| xlPyramidColStacked | 107 | Stacked pyramid column |
| xlPyramidColStacked100 | 108 | 100% stacked pyramid column |
| xlRadar | -4151 | Radar |
| xlRadarFilled | 82 | Filled radar |
| xlRadarMarkers | 81 | Radar with markers |
| xlStockHLC | 88 | High-low-close stock |
| xlStockOHLC | 89 | Open-high-low-close stock |
| xlStockVHLC | 90 | Volume-high-low-close stock |
| xlStockVOHLC | 91 | Volume-open-high-low-close stock |
| xlSurface | 83 | 3-D surface |
| xlSurfaceTopView | 85 | Surface top view |
| xlSurfaceTopViewWireframe | 86 | Surface top view wireframe |
| xlSurfaceWireframe | 84 | 3-D surface wireframe |
| xlXYScatter | -4169 | Scatter |
| xlXYScatterLines | 74 | Scatter with lines |
| xlXYScatterLinesNoMarkers | 75 | Scatter with lines, no markers |
| xlXYScatterSmooth | 72 | Scatter with smooth lines |
| xlXYScatterSmoothNoMarkers | 73 | Scatter with smooth lines, no markers |

## XlContainsOperator

Operators used by contains functions.

| Field | Value | Meaning |
|---|---|---|
| xlContains | 0 | Contains the value |
| xlDoesNotContain | 1 | Does not contain the value |
| xlBeginsWith | 2 | Begins with the value |
| xlEndsWith | 3 | Ends with the value |

## XlDeleteShiftDirection

How cells shift when deleting.

| Field | Value | Meaning |
|---|---|---|
| xlShiftToLeft | -4159 | Shift cells left |
| xlShiftUp | -4162 | Shift cells up |

## XlDirection

Direction to move.

| Field | Value | Meaning |
|---|---|---|
| xlDown | -4121 | Down |
| xlToLeft | -4159 | Left |
| xlToRight | -4161 | Right |
| xlUp | -4162 | Up |

## XlDVAlertStyle

Icon used in validation alerts.

| Field | Value | Meaning |
|---|---|---|
| xlValidAlertStop | 1 | Stop icon |
| xlValidAlertWarning | 2 | Warning icon |
| xlValidAlertInformation | 3 | Information icon |

## XlDVType

Type of validation to perform on values.

| Field | Value | Meaning |
|---|---|---|
| xlValidateWholeNumber | 1 | Whole number |
| xlValidateDecimal | 2 | Decimal |
| xlValidateList | 3 | Value must exist in list |
| xlValidateDate | 4 | Date |
| xlValidateTime | 5 | Time |
| xlValidateTextLength | 6 | Text length |
| xlValidateCustom | 7 | Custom formula validation |

## XlExportImgFormatType

Exported image format.

| Field | Value | Meaning |
|---|---|---|
| xlImgTypePNG | 0 | Export .png |
| xlImgTypeJPG | 1 | Export .jpg |
| xlImgTypeBMP | 2 | Export .bmp |
| xlImgTypeTIF | 2 | Export .tif |

## XlFixedFormatType

Fixed export format type.

| Field | Value | Meaning |
|---|---|---|
| xlTypePDF | 0 | PDF (.pdf) |
| xlTypeXPS | 1 | XPS (.xps) |
| xlTypeIMG | 2 | IMG (.png, .jpg, .bmp, .tif) |

## XlFormatConditionOperator

Operator comparing formulas to cell values, including xlBetween and xlNotBetween.

| Field | Value | Meaning |
|---|---|---|
| xlBetween | 1 | Between; requires two formulas |
| xlNotBetween | 2 | Not between; requires two formulas |
| xlEqual | 3 | Equal |
| xlNotEqual | 4 | Not equal |
| xlGreater | 5 | Greater than |
| xlLess | 6 | Less than |
| xlGreaterEqual | 7 | Greater than or equal |
| xlLessEqual | 8 | Less than or equal |

## XlFormatConditionType

Whether conditional format is based on cell values or expressions.

| Field | Value | Meaning |
|---|---|---|
| xlCellValue | 1 | Cell value |
| xlExpression | 2 | Expression |
| xlColorScale | 3 | Color scale |
| xlTop10 | 5 | Top or bottom values |
| xlUniqueValues | 8 | Unique values |
| xlTextString | 9 | Text string |
| xlBlanksCondition | 10 | Blanks condition |
| xlTimePeriod | 11 | Time period |
| xlAboveAverageCondition | 12 | Above average condition |
| xlNoBlanksCondition | 13 | No blanks condition |
| xlErrorsCondition | 16 | Errors condition |
| xlNoErrorsCondition | 17 | No errors condition |

## XlHAlign

Horizontal alignment.

| Field | Value | Meaning |
|---|---|---|
| xlHAlignRight | -4152 | Right |
| xlHAlignLeft | -4131 | Left |
| xlHAlignJustify | -4130 | Justify |
| xlHAlignDistributed | -4117 | Distributed |
| xlHAlignCenter | -4108 | Center |
| xlHAlignGeneral | 1 | General (by data type) |
| xlHAlignFill | 5 | Fill |
| xlHAlignCenterAcrossSelection | 7 | Center across selection |

## XlInsertFormatOrigin

Where to copy inserted cell formats from.

| Field | Value | Meaning |
|---|---|---|
| xlFormatFromLeftOrAbove | 0 | Copy format from left or above |
| xlFormatFromRightOrBelow | 1 | Copy format from right or below |

## XlInsertShiftDirection

Direction cells shift when inserting.

| Field | Value | Meaning |
|---|---|---|
| xlShiftDown | -4121 | Shift cells down |
| xlShiftToRight | -4161 | Shift cells right |

## XlLineStyle

Border line style.

| Field | Value | Meaning |
|---|---|---|
| xlLineStyleNone | -4142 | No line |
| xlDouble | -4119 | Double line |
| xlDot | -4118 | Dotted |
| xlDash | -4115 | Dashed |
| xlContinuous | 1 | Solid |
| xlDashDot | 4 | Dash-dot |
| xlDashDotDot | 5 | Dash dot-dot |
| xlSlantDashDot | 13 | Slant dash-dot |

## XlPasteSpecialOperation

Paste special operation.

| Field | Value | Meaning |
|---|---|---|
| xlPasteSpecialOperationAdd | 0x1 | Add pasted data to destination |
| xlPasteSpecialOperationDivide | 0x4 | Divide pasted data by destination |
| xlPasteSpecialOperationMultiply | 0x3 | Multiply pasted data by destination |
| xlPasteSpecialOperationNone | 0x0 | No operation during paste |
| xlPasteSpecialOperationSubtract | 0x2 | Subtract destination from pasted data |

## XlPasteType

Paste type.

| Field | Value | Meaning |
|---|---|---|
| xlPasteAll | 0x1 | Paste all |
| xlPasteFormulas | 0x2 | Paste formulas |
| xlPasteAllExceptBorders | 0x4 | Paste all except borders |
| xlPasteColumnWidths | 0x5 | Paste column widths |
| xlPasteValues | 0x3 | Paste values |
| xlPasteValuesAndNumberFormats | 0x7 | Paste values and number formats |
| xlPasteFormats | 0x8 | Paste formats |
| xlPasteComments | 0x9 | Paste comments |
| xlPasteValidation | 0xa | Paste validation |

## XlReferenceStyle

Reference style.

| Field | Value | Meaning |
|---|---|---|
| xlR1C1 | -4150 | Use xlR1C1 for R1C1-style references |
| xlA1 | 1 | Default. Use xlA1 for A1-style references |

## XlRowCol

Whether series values are organized by rows or columns.

| Field | Value | Meaning |
|---|---|---|
| xlRows | 1 | Data series in rows |
| xlColumns | 2 | Data series in columns |

## XlSheetType

Worksheet type.

| Field | Value | Meaning |
|---|---|---|
| xlWorksheet | -4167 | Worksheet |
| xlDialogSheet | -4116 | Dialog sheet |
| xlChart | -4109 | Chart |
| xlExcel4MacroSheet | 3 | Excel 4.0 macro sheet |
| xlExcel4IntlMacroSheet | 4 | Excel 4.0 international macro sheet |

## XlTimePeriods

Time period filter.

| Field | Value | Meaning |
|---|---|---|
| xlToday | 0 | Today |
| xlYesterday | 1 | Yesterday |
| xlLast7Days | 2 | Last 7 days |
| xlThisWeek | 3 | This week |
| xlLastWeek | 4 | Last week |
| xlLastMonth | 5 | Last month |
| xlTomorrow | 6 | Tomorrow |
| xlNextWeek | 7 | Next week |
| xlNextMonth | 8 | Next month |
| xlThisMonth | 9 | This month |

## XlUnderlineStyle

Font underline style.

| Field | Value | Meaning |
|---|---|---|
| xlUnderlineStyleDouble | -4119 | Double underline |
| xlUnderlineStyleDoubleAccounting | 5 | Double accounting underline |
| xlUnderlineStyleNone | -4142 | No underline |
| xlUnderlineStyleSingle | 2 | Single underline |
| xlUnderlineStyleSingleAccounting | 4 | Not supported |

## XlVAlign

Vertical alignment.

| Field | Value | Meaning |
|---|---|---|
| xlVAlignTop | -4160 | Top |
| xlVAlignJustify | -4130 | Justify (vertical) |
| xlVAlignDistributed | -4117 | Distributed (vertical) |
| xlVAlignCenter | -4108 | Center |
| xlVAlignBottom | -4107 | Bottom |

## XlXLMMacroType

Macro type referenced by a name on the sheet.

| Field | Value | Meaning |
|---|---|---|
| xlFunction | 1 | Custom function |
| xlCommand | 2 | Custom command |
| xlNotXLM | 3 | Not a macro |

## XlAutoFilterOperator

AutoFilter operator or logical combination of two criteria.

| Field | Value | Meaning |
|---|---|---|
| xlAnd | 1 | Logical AND of Criteria1 and Criteria2 |
| xlOr | 2 | Logical OR of Criteria1 and Criteria2 |
| xlTop10Items | 3 | Top N items (N in Criteria1) |
| xlBottom10Items | 4 | Bottom N items (N in Criteria1) |
| xlTop10Percent | 5 | Top percent (percent in Criteria1) |
| xlBottom10Percent | 6 | Bottom percent (percent in Criteria1) |
| xlFilterValues | 7 | Filter by values |
| xlFilterCellColor | 8 | Cell color |
| xlFilterFontColor | 9 | Font color |
| xlFilterIcon | 10 | Filter by icon |
| xlFilterDynamic | 11 | Dynamic filter |

## XlYesNoGuess

Whether the first row contains headers. Not used when sorting PivotTables.

| Field | Value | Meaning |
|---|---|---|
| xlGuess | 0 | Let Excel detect headers |
| xlYes | 1 | Default. Sort the entire range |
| xlNo | 2 | Do not sort entire range as headers |

## XlSortOrientation

Sort orientation (rows vs columns).

| Field | Value | Meaning |
|---|---|---|
| xlSortColumns | 1 | Sort by columns |
| xlSortRows | 2 | Sort by rows (default) |

## XlSortMethod

East Asian sort method.

| Field | Value | Meaning |
|---|---|---|
| xlPinYin | 1 | Phonetic order for East Asian text (default) |
| xlStroke | 2 | Stroke count order for East Asian text |

## XlSortDataOption

How text is sorted.

| Field | Value | Meaning |
|---|---|---|
| xlSortNormal | 0 | Sort numbers and text separately (default) |
| xlSortTextAsNumbers | 1 | Sort text as numbers |

## XlSortOn

Property to sort on.

| Field | Value | Meaning |
|---|---|---|
| SortOnCellColor | 1 | Cell color |
| SortOnFontColor | 2 | Font color |
| SortOnIcon | 3 | Icon |
| SortOnValues | 0 | Values |

## XlSortOrder

Sort order for a field or range.

| Field | Value | Meaning |
|---|---|---|
| xlAscending | 1 | Default. Ascending sort for the field |
| xlDescending | 2 | Descending sort |

## XlTextParsingType

Column format when importing delimited text.

| Field | Value | Meaning |
|---|---|---|
| xlDelimited | 1 | Default. Delimited text file |
| xlFixedWidth | 2 | Fixed-width columns |

## XlTextQualifier

Text qualifier for parsing.

| Field | Value | Meaning |
|---|---|---|
| xlTextQualifierDoubleQuote | 1 | Double quote (") |
| xlTextQualifierNone | -4142 | No qualifier |
| xlTextQualifierSingleQuote | 2 | Single quote (') |

## XlColumnDataType

Column data type when splitting text.

| Field | Value | Meaning |
|---|---|---|
| xlDMYFormat | 4 | DMY date format |
| xlDYMFormat | 7 | DYM date format |
| xlEMDFormat | 10 | EMD date format |
| xlGeneralFormat | 1 | General |
| xlMDYFormat | 3 | MDY date format |
| xlMYDFormat | 6 | MYD date format |
| xlSkipColumn | 9 | Do not import column |
| xlTextFormat | 2 | Text |
| xlYDMFormat | 8 | YDM date format |
| xlYMDFormat | 5 | YMD date format |
