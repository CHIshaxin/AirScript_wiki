---
title: "Enum"
source: https://airsheet.wps.cn/docs/en-US/api/excel/databook/Enum.html
updated: 1789700916000
hash: DGAhu11K
---

# Enum

Enumeration types are exposed on `Application`.

## XlAboveBelow

Specifies whether a value is above or below the average.

| Member | Value | Description |
| --- | --- | --- |
| xlAboveAverage | 0 | Above average |
| xlBelowAverage | 1 | Below average |
| xlEqualAboveAverage | 2 | Equal to average |

## XlAutoFillType

Specifies how the destination range is filled based on the source range.

| Member | Value | Description |
| --- | --- | --- |
| xlFillDefault | 0 | Determine values and formats used to fill the destination |
| xlFillCopy | 1 | Copy values and formats from source to destination, repeating if needed |
| xlFillSeries | 2 | Extend values from source to destination as a series (e.g. “1, 2” becomes “3, 4, 5”). Copy formats from source, repeating if needed |
| xlFillFormats | 3 | Copy only formats from source to destination, repeating if needed |
| xlFillValues | 4 | Copy only values from source to destination, repeating if needed |
| xlFillDays | 5 | Extend weekday names from source to destination. Copy formats from source, repeating if needed |
| xlFillWeekdays | 6 | Extend workweek day names from source to destination. Copy formats from source, repeating if needed |
| xlFillMonths | 7 | Extend month names from source to destination. Copy formats from source, repeating if needed |
| xlFillYears | 8 | Extend years from source to destination. Copy formats from source, repeating if needed |
| xlLinearTrend | 9 | Extend numbers assuming an additive trend (e.g. “1, 2,” to “3, 4, 5”). Copy formats from source, repeating if needed |
| xlGrowthTrend | 10 | Extend numbers assuming a multiplicative trend (e.g. “1, 2,” to “4, 8, 16”). Copy formats from source, repeating if needed |

## XlBordersIndex

Specifies which border to retrieve.

| Member | Value | Description |
| --- | --- | --- |
| xlDiagonalDown | 5 | Border from top-left to bottom-right of each cell |
| xlDiagonalUp | 6 | Border from bottom-left to top-right of each cell |
| xlEdgeLeft | 7 | Left edge border of the range |
| xlEdgeTop | 8 | Top border of the range |
| xlEdgeBottom | 9 | Bottom border of the range |
| xlEdgeRight | 10 | Right edge border of the range |
| xlInsideVertical | 11 | Vertical borders inside the range (excluding outside edges) |
| xlInsideHorizontal | 12 | Horizontal borders inside the range (excluding outside edges) |
| xlOutside | 13 | Top, bottom, left, and right of the range |
| xlInside | 14 | Interior region |

## XlBorderWeight

Specifies border weight around a range.

| Member | Value | Description |
| --- | --- | --- |
| xlMedium | -4138 | Medium |
| xlHairline | 1 | Hairline (thinnest border) |
| xlThin | 2 | Thin |
| xlThick | 4 | Thick (widest border) |

## XlCalcModeType

Iterative calculation mode.

| Member | Value | Description |
| --- | --- | --- |
| manual | `manual` | Manual |
| automatic | `automatic` | Automatic |

## XlChartType

Specifies the chart type.

| Member | Value | Description |
| --- | --- | --- |
| xl3DArea | -4098 | 3-D area chart. |
| xl3DAreaStacked | 78 | 3-D stacked area chart. |
| xl3DAreaStacked100 | 79 | 100% stacked area chart. |
| xl3DBarClustered | 60 | 3-D clustered bar chart. |
| xl3DBarStacked | 61 | 3-D stacked bar chart. |
| xl3DBarStacked100 | 62 | 3-D 100% stacked bar chart. |
| xl3DColumn | -4100 | 3-D column chart. |
| xl3DColumnClustered | 54 | 3-D clustered column chart. |
| xl3DColumnStacked | 55 | 3-D stacked column chart. |
| xl3DColumnStacked100 | 56 | 3-D 100% stacked column chart. |
| xl3DLine | -4101 | 3-D line chart. |
| xl3DPie | -4102 | 3-D pie chart. |
| xl3DPieExploded | 70 | Exploded 3-D pie chart. |
| xlArea | 1 | Area chart |
| xlAreaStacked | 76 | Stacked area chart. |
| xlAreaStacked100 | 77 | 100% stacked area chart. |
| xlBarClustered | 57 | Clustered bar chart. |
| xlBarOfPie | 71 | Bar of pie chart. |
| xlBarStacked | 58 | Stacked bar chart. |
| xlBarStacked100 | 59 | 100% stacked bar chart. |
| xlBubble | 15 | Bubble chart. |
| xlBubble3DEffect | 87 | 3-D bubble chart. |
| xlColumnClustered | 51 | Clustered column chart. |
| xlColumnStacked | 52 | Stacked column chart. |
| xlColumnStacked100 | 53 | 100% stacked column chart. |
| xlConeBarClustered | 102 | Clustered cone bar chart. |
| xlConeBarStacked | 103 | Stacked cone bar chart. |
| xlConeBarStacked100 | 104 | 100% stacked cone bar chart. |
| xlConeCol | 105 | 3-D cone column chart. |
| xlConeColClustered | 99 | Clustered cone column chart. |
| xlConeColStacked | 100 | Stacked cone column chart. |
| xlConeColStacked100 | 101 | 100% stacked cone column chart. |
| xlCylinderBarClustered | 95 | Clustered cylinder bar chart. |
| xlCylinderBarStacked | 96 | Stacked cylinder bar chart. |
| xlCylinderBarStacked100 | 97 | 100% stacked cylinder bar chart. |
| xlCylinderCol | 98 | 3-D cylinder column chart. |
| xlCylinderColClustered | 92 | Clustered cone column chart. |
| xlCylinderColStacked | 93 | Stacked cone column chart. |
| xlCylinderColStacked100 | 94 | 100% stacked cylinder column chart. |
| xlDoughnut | -4120 | Doughnut chart. |
| xlDoughnutExploded | 80 | Exploded doughnut chart. |
| xlLine | 4 | Line chart. |
| xlLineMarkers | 65 | Line chart with markers. |
| xlLineMarkersStacked | 66 | Stacked line chart with markers. |
| xlLineMarkersStacked100 | 67 | 100% stacked line chart with markers. |
| xlLineStacked | 63 | Stacked line chart. |
| xlLineStacked100 | 64 | 100% stacked line chart. |
| xlPie | 5 | Pie chart. |
| xlPieExploded | 69 | Exploded pie chart. |
| xlPieOfPie | 68 | Pie of pie chart. |
| xlPyramidBarClustered | 109 | Clustered pyramid bar chart. |
| xlPyramidBarStacked | 110 | Stacked pyramid bar chart. |
| xlPyramidBarStacked100 | 111 | 100% stacked pyramid bar chart. |
| xlPyramidCol | 112 | 3-D pyramid column chart. |
| xlPyramidColClustered | 106 | Clustered pyramid column chart. |
| xlPyramidColStacked | 107 | Stacked pyramid column chart. |
| xlPyramidColStacked100 | 108 | 100% stacked pyramid column chart. |
| xlRadar | -4151 | Radar chart. |
| xlRadarFilled | 82 | Filled radar chart. |
| xlRadarMarkers | 81 | Radar chart with markers. |
| xlStockHLC | 88 | High-low-close stock chart. |
| xlStockOHLC | 89 | Open-high-low-close stock chart. |
| xlStockVHLC | 90 | Volume-high-low-close stock chart. |
| xlStockVOHLC | 91 | Volume-open-high-low-close stock chart. |
| xlSurface | 83 | 3-D surface chart. |
| xlSurfaceTopView | 85 | Surface chart (top view). |
| xlSurfaceTopViewWireframe | 86 | Surface chart top view (wireframe). |
| xlSurfaceWireframe | 84 | 3-D surface chart (wireframe). |
| xlXYScatter | -4169 | Scatter chart. |
| xlXYScatterLines | 74 | Scatter chart with straight lines. |
| xlXYScatterLinesNoMarkers | 75 | Scatter chart with lines, no markers. |
| xlXYScatterSmooth | 72 | Scatter chart with smooth lines. |
| xlXYScatterSmoothNoMarkers | 73 | Scatter chart with smooth lines, no markers. |

## XlContainsOperator

Specifies operators used by the function.

| Member | Value | Description |
| --- | --- | --- |
| xlContains | 0 | Contains the specified value |
| xlDoesNotContain | 1 | Does not contain the specified value |
| xlBeginsWith | 2 | Begins with the specified value |
| xlEndsWith | 3 | Ends with the specified value |

## XlDeleteShiftDirection

Specifies how cells move to replace deleted cells.

| Member | Value | Description |
| --- | --- | --- |
| xlShiftToLeft | -4159 | Shift cells left |
| xlShiftUp | -4162 | Shift cells up |

## XlDirection

Specifies the direction to move.

| Member | Value | Description |
| --- | --- | --- |
| xlDown | -4121 | Down |
| xlToLeft | -4159 | To the left |
| xlToRight | -4161 | To the right |
| xlUp | -4162 | Up |

## XlDVAlertStyle

Specifies the icon used in validation message boxes.

| Member | Value | Description |
| --- | --- | --- |
| xlValidAlertStop | 1 | Stop icon |
| xlValidAlertWarning | 2 | Warning icon |
| xlValidAlertInformation | 3 | Information icon |

## XlDVType

Specifies the type of validation to apply to values.

| Member | Value | Description |
| --- | --- | --- |
| xlValidateWholeNumber | 1 | Whole number |
| xlValidateDecimal | 2 | Decimal |
| xlValidateList | 3 | Value must exist in the specified list |
| xlValidateDate | 4 | Date |
| xlValidateTime | 5 | Time |
| xlValidateTextLength | 6 | Text length |
| xlValidateCustom | 7 | Custom formula validation |

## XlExportImgFormatType

Image export format.

| Member | Value | Description |
| --- | --- | --- |
| xlImgTypePNG | 0 | Export .png |
| xlImgTypeJPG | 1 | Export .jpg |
| xlImgTypeBMP | 2 | Export .bmp |
| xlImgTypeTIF | 2 | Export .tif |

## XlFixedFormatType

Specifies the file format type.

| Member | Value | Description |
| --- | --- | --- |
| xlTypePDF | 0 | PDF (.pdf) |
| xlTypeXPS | 1 | XPS (.xps) |
| xlTypeIMG | 2 | IMG (.png, .jpg, .bmp, .tif) |

## XlFormatConditionOperator

Specifies operators that compare formulas to cell values, or compare two formulas for xlBetween and xlNotBetween.

| Member | Value | Description |
| --- | --- | --- |
| xlBetween | 1 | Between; only valid when two formulas are provided |
| xlNotBetween | 2 | Not between; only valid when two formulas are provided |
| xlEqual | 3 | Equal |
| xlNotEqual | 4 | Not equal |
| xlGreater | 5 | Greater than |
| xlLess | 6 | Less than |
| xlGreaterEqual | 7 | Greater than or equal to |
| xlLessEqual | 8 | Less than or equal to |

## XlFormatConditionType

Specifies whether conditional formatting is based on cell values or expressions.

| Member | Value | Description |
| --- | --- | --- |
| xlCellValue | 1 | Cell value |
| xlExpression | 2 | Expression |
| xlColorScale | 3 | Color scale |
| xlTop10 | 5 | Top 10 values |
| xlUniqueValues | 8 | Unique values |
| xlTextString | 9 | Text string |
| xlBlanksCondition | 10 | Blanks condition |
| xlTimePeriod | 11 | Time period |
| xlAboveAverageCondition | 12 | Above-average condition |
| xlNoBlanksCondition | 13 | No-blanks condition |
| xlErrorsCondition | 16 | Errors condition |
| xlNoErrorsCondition | 17 | No-errors condition |

## XlHAlign

Specifies horizontal alignment.

| Member | Value | Description |
| --- | --- | --- |
| xlHAlignRight | -4152 | Right |
| xlHAlignLeft | -4131 | Left |
| xlHAlignJustify | -4130 | Justify |
| xlHAlignDistributed | -4117 | Distributed |
| xlHAlignCenter | -4108 | Center |
| xlHAlignGeneral | 1 | General (by data type) |
| xlHAlignFill | 5 | Fill |
| xlHAlignCenterAcrossSelection | 7 | Center across selection |

## XlInsertFormatOrigin

Specifies where to copy formats when inserting cells.

| Member | Value | Description |
| --- | --- | --- |
| xlFormatFromLeftOrAbove | 0 | Copy format from left or above |
| xlFormatFromRightOrBelow | 1 | Copy format from right or below |

## XlInsertShiftDirection

Specifies how cells shift when inserting.

| Member | Value | Description |
| --- | --- | --- |
| xlShiftDown | -4121 | Shift cells down |
| xlShiftToRight | -4161 | Shift cells to the right |

## XlLineStyle

Specifies the line style for borders.

| Member | Value | Description |
| --- | --- | --- |
| xlLineStyleNone | -4142 | None |
| xlDouble | -4119 | Double |
| xlDot | -4118 | Dotted |
| xlDash | -4115 | Dashed |
| xlContinuous | 1 | Continuous |
| xlDashDot | 4 | Dash-dot |
| xlDashDotDot | 5 | Dash dot-dot |
| xlSlantDashDot | 13 | Slant dash-dot |

## XlPasteSpecialOperation

Special operations for XlPaste.

| Member | Value | Description |
| --- | --- | --- |
| xlPasteSpecialOperationAdd | 0x1 | Add copied data to destination values |
| xlPasteSpecialOperationDivide | 0x4 | Divide copied data by destination values |
| xlPasteSpecialOperationMultiply | 0x3 | Multiply copied data by destination values |
| xlPasteSpecialOperationNone | 0x0 | No calculation during paste |
| xlPasteSpecialOperationSubtract | 0x2 | Subtract destination values from copied data |

## XlPasteType

Paste type.

| Member | Value | Description |
| --- | --- | --- |
| xlPasteFormulas | 0x2 | Formulas |
| xlPasteAllExceptBorders | 0x4 | All except borders |
| xlPasteColumnWidths | 0x5 | Column widths |
| xlPasteValues | 0x3 | Values |
| xlPasteValuesAndNumberFormats | 0x7 | Values and number formats |
| xlPasteFormats | 0x8 | Formats |
| xlPastePasteAll | 0x0 | Paste all |
| xlPasteComments | 0x9 | Comments |
| xlPasteValidation | 0xa | Validation |

## XlReferenceStyle

Specifies the reference style.

| Member | Value | Description |
| --- | --- | --- |
| xlR1C1 | -4150 | Use xlR1C1 to return R1C1-style references |
| xlA1 | 1 | Default. Use xlA1 to return A1-style references |

## XlRowCol

Specifies whether data series values are organized in rows or columns.

| Member | Value | Description |
| --- | --- | --- |
| xlRows | 1 | Data series in columns |
| xlColumns | 2 | Data series in rows |

## XlSheetType

Specifies the worksheet type.

| Member | Value | Description |
| --- | --- | --- |
| xlWorksheet | -4167 | Worksheet |
| xlDialogSheet | -4116 | Dialog sheet |
| xlChart | -4109 | Chart |
| xlExcel4MacroSheet | 3 | Excel 4.0 macro sheet |
| xlExcel4IntlMacroSheet | 4 | Excel 4.0 international macro sheet |

## XlTimePeriods

Specifies the time period.

| Member | Value | Description |
| --- | --- | --- |
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

Specifies the underline style applied to the font.

| Member | Value | Description |
| --- | --- | --- |
| xlUnderlineStyleDouble | -4119 | Double underline |
| xlUnderlineStyleDoubleAccounting | 5 | Double accounting (two thin underlines) |
| xlUnderlineStyleNone | -4142 | No underline |
| xlUnderlineStyleSingle | 2 | Single underline |
| xlUnderlineStyleSingleAccounting | 4 | Not supported |

## XlVAlign

Specifies vertical alignment.

| Member | Value | Description |
| --- | --- | --- |
| xlVAlignTop | -4160 | Top |
| xlVAlignJustify | -4130 | Justify (vertical) |
| xlVAlignDistributed | -4117 | Distributed (vertical) |
| xlVAlignCenter | -4108 | Center |
| xlVAlignBottom | -4107 | Bottom |

## XlXLMMacroType

Specifies whether a name refers to a macro in the worksheet, or whether it refers to a macro at all.

| Member | Value | Description |
| --- | --- | --- |
| xlFunction | 1 | User-defined function |
| xlCommand | 2 | User-defined command |
| xlNotXLM | 3 | Not a macro |

## XlAutoFilterOperator

Specifies the filter type on the current worksheet, or the operator that joins two filter criteria.

| Member | Value | Description |
| --- | --- | --- |
| xlAnd | 1 | Logical AND of Criteria1 and Criteria2 |
| xlOr | 2 | Logical OR of Criteria1 and Criteria2 |
| xlTop10Items | 3 | Top items (count in Criteria1) |
| xlBottom10Items | 4 | Bottom items (count in Criteria1) |
| xlTop10Percent | 5 | Top percent (percent in Criteria1) |
| xlBottom10Percent | 6 | Bottom percent (percent in Criteria1) |
| xlFilterValues | 7 | Filter values |
| xlFilterCellColor | 8 | Cell color |
| xlFilterFontColor | 9 | Font color |
| xlFilterIcon | 10 | Filter icon |
| xlFilterDynamic | 11 | Dynamic filter |

## XlYesNoGuess

Specifies whether the first row contains headers. Cannot be used when sorting a PivotTable.

| Member | Value | Description |
| --- | --- | --- |
| xlGuess | 0 | Guess whether Excel has headers |
| xlYes | 1 | Default. Sort the entire range |
| xlNo | 2 | Do not sort the entire range |

## XlSortOrientation

Specifies sort orientation.

| Member | Value | Description |
| --- | --- | --- |
| xlSortColumns | 1 | Sort by columns |
| xlSortRows | 2 | Sort by rows (default) |

## XlSortMethod

Specifies the sort method.

| Member | Value | Description |
| --- | --- | --- |
| xlPinYin | 1 | Sort by Pinyin for Chinese characters (default) |
| xlStroke | 2 | Sort by stroke count for each character |

## XlSortDataOption

Specifies how text is sorted.

| Member | Value | Description |
| --- | --- | --- |
| xlSortNormal | 0 | Sort numbers and text separately (default) |
| xlSortTextAsNumbers | 1 | Sort text as numeric data |

## XlSortOn

Specifies what to sort on.

| Member | Value | Description |
| --- | --- | --- |
| SortOnCellColor | 1 | Cell color |
| SortOnFontColor | 2 | Font color |
| SortOnIcon | 3 | Icon |
| SortOnValues | 0 | Values |

## XlSortOrder

Specifies sort order for a field or range.

| Member | Value | Description |
| --- | --- | --- |
| xlAscending | 1 | Default. Sort ascending |
| xlDescending | 2 | Sort descending |

## XlTextParsingType

Specifies the column data format when importing a text file into a query table.

| Member | Value | Description |
| --- | --- | --- |
| xlDelimited | 1 | Default. Delimited file |
| xlFixedWidth | 2 | Fixed-width columns |

## XlTextQualifier

Specifies the text qualifier.

| Member | Value | Description |
| --- | --- | --- |
| xlTextQualifierDoubleQuote | 1 | Double quote (") |
| xlTextQualifierNone | -4142 | No qualifier |
| xlTextQualifierSingleQuote | 2 | Single quote (') |

## XlColumnDataType

Specifies how columns are parsed.

| Member | Value | Description |
| --- | --- | --- |
| xlDMYFormat | 4 | DMY date |
| xlDYMFormat | 7 | DYM date |
| xlEMDFormat | 10 | EMD date |
| xlGeneralFormat | 1 | General |
| xlMDYFormat | 3 | MDY date |
| xlMYDFormat | 6 | MYD date |
| xlSkipColumn | 9 | Skip column |
| xlTextFormat | 2 | Text |
| xlYDMFormat | 8 | YDM date |
| xlYMDFormat | 5 | YMD date |
