---
title: "ClearLabelFilters (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E6%96%B9%E6%B3%95/ClearLabelFilters%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cb3wdgJC
---

# ClearLabelFilters (方法)

此方法将删除透视字段的 **[PivotFilters](./../../PivotFilters/PivotFilters 对象.html)** 集合中的所有标签筛选或所有日期筛选。

## 说明

下表列出了此方法将删除的各种标签筛选类型。

| **xlCaptionEquals** |
| --- |
| **xlCaptionDoesNotEqual** |
| **xlCaptionIsGreaterThan** |
| **xlCaptionIsGreaterThanOrEqualTo** |
| **xlCaptionIsLessThan** |
| **xlCaptionIsLessThanOrEqualTo** |
| **xlCaptionBeginsWith** |
| **xlCaptionDoesNotBeginWith** |
| **xlCaptionEndsWith** |
| **xlCaptionDoesNotEndWith** |
| **xlCaptionContains** |
| **xlCaptionDoesNotContain** |
| **xlCaptionIsBetween** |
| **xlCaptionIsNotBetween** |

下表列出了此方法将删除的各种日期筛选类型。

| **xlSpecificDate** |
| --- |
| **xlNotSpecificDate** |
| **xlBefore** |
| **xlBeforeOrEqualTo** |
| **xlAfter** |
| **xlAfterOrEqualTo** |
| **xlDateBetween** |
| **xlDateNotBetween** |
| **xlDateToday** |
| **xlDateYesterday** |
| **xlDateTomorrow** |
| **xlDateThisWeek** |
| **xlDateLastWeek** |
| **xlDateNextWeek** |
| **xlDateThisMonth** |
| **xlDateLastMonth** |
| **xlDateNextMonth** |
| **xlDateThisQuarter** |
| **xlDateLastQuarter** |
| **xlDateNextQuarter** |
| **xlDateThisYear** |
| **xlDateLastYear** |
| **xlDateNextYear** |
| **xlYearToDate** |
| **xlAllDatesInPeriodQuarter1** |
| **xlAllDatesInPeriodQuarter2** |
| **xlAllDatesInPeriodQuarter3** |
| **xlAllDatesInPeriodQuarter4** |
| **xlAllDatesInPeriodJanuary** |
| **xlAllDatesInPeriodFebruary** |
| **xlAllDatesInPeriodMarch** |
| **xlAllDatesInPeriodApril** |
| **xlAllDatesInPeriodMay** |
| **xlAllDatesInPeriodJune** |
| **xlAllDatesInPeriodJuly** |
| **xlAllDatesInPeriodAugust** |
| **xlAllDatesInPeriodSeptember** |
| **xlAllDatesInPeriodOctober** |
| **xlAllDatesInPeriodNovember** |
| **xlAllDatesInPeriodDecember** |

## 示例

```javascript
/*本示例删除活动工作表中第一张数据透视表的字段“id”的所有日期筛选器。*/
function test() {
    ActiveSheet.PivotTables(1).PivotFields("id").ClearAllFilters()
}
```

```javascript
/*本示例删除工作表 Sheet1 中数据透视表应用于字段“grades”的所有标签筛选器。*/
function test() {
    let pvtfield = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("grades")
    pvtfield.ClearLabelFilters()
}
```
