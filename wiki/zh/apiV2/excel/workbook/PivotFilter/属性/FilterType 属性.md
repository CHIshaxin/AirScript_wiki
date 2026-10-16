---
title: "FilterType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/FilterType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BmBNgQS7
---

# FilterType (属性)

指定要应用的筛选器类型。只读 **xlPivotFilterType** 类型。

## 说明

下表列出了新的筛选器类型，以及对于每一个筛选器类型必须提供的参数和不可提供（不可用）的参数。

| xlPivotFilterType | DataField | DataCubeField (OLAP) | Value1 | Value2 |
| --- | --- | --- | --- | --- |
| xlTopCount | 必需 | 必需 | 必需 | 不可用 |
| xlBottomCount | 必需 | 必需 | 必需 | 不可用 |
| xlTopPercent | 必需 | 必需 | 必需 | 不可用 |
| xlBottomPercent | 必需 | 必需 | 必需 | 不可用 |
| xlTopSum | 必需 | 必需 | 必需 | 不可用 |
| xlBottomSum | 必需 | 必需 | 必需 | 不可用 |
| xlCaptionEquals | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionDoesNotEqual | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionIsGreaterThan | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionIsGreaterThanOrEqualTo | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionIsLessThan | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionIsLessThanOrEqualTo | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionBeginsWith | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionDoesNotBeginWith | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionEndsWith | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionDoesNotEndWith | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionContains | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionDoesNotContain | 不可用 | 不可用 | 必需 | 不可用 |
| xlCaptionIsBetween | 不可用 | 不可用 | 必需 | 必需 |
| xlCaptionIsNotBetween | 不可用 | 不可用 | 必需 | 必需 |
| xlValueEquals | 必需 | 必需 | 必需 | 不可用 |
| xlValueDoesNotEqual | 必需 | 必需 | 必需 | 不可用 |
| xlValueIsGreaterThan | 必需 | 必需 | 必需 | 不可用 |
| xlValueIsGreaterThanOrEqualTo | 必需 | 必需 | 必需 | 不可用 |
| xlValueIsLessThan | 必需 | 必需 | 必需 | 不可用 |
| xlValueIsLessThanOrEqualTo | 必需 | 必需 | 必需 | 不可用 |
| xlValueIsBetween | 必需 | 必需 | 必需 | 必需 |
| xlValueIsNotBetween | 必需 | 必需 | 必需 | 必需 |
| xlSpecificDate | 不可用 | 不可用 | 必需 | 不可用 |
| xlNotSpecificDate | 不可用 | 不可用 | 必需 | 不可用 |
| xlBefore | 不可用 | 不可用 | 必需 | 不可用 |
| xlBeforeOrEqualTo | 不可用 | 不可用 | 必需 | 不可用 |
| xlAfter | 不可用 | 不可用 | 必需 | 不可用 |
| xlAfterOrEqualTo | 不可用 | 不可用 | 必需 | 不可用 |
| xlBetween | 不可用 | 不可用 | 必需 | 不可用 |
| xlNotBetween | 不可用 | 不可用 | 必需 | 不可用 |
| xlFilterToday | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterYesterday | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterTomorrow | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterThisWeek | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterLastWeek | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterNextWeek | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterThisMonth | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterLastMonth | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterNextMonth | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterThisQuarter | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterLastQuarter | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterNextQuarter | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterThisYear | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterLastYear | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterNextYear | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterYearToDate | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter1 | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter2 | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter3 | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodQuarter4 | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodJanuary | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodFebruary | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodMarch | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodApril | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodMay | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodJune | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodJuly | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodAugust | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodSeptember | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodOctober | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodNovember | 不可用 | 不可用 | 不可用 | 不可用 |
| xlFilterAllDatesInPeriodDecember | 不可用 | 不可用 | 不可用 | 不可用 |

## 示例

```javascript
/*本示例显示工作表 Sheet1 中数据透视表的字段“score”的第一个筛选器应用的筛选器类型是否为“xlCaptionIsGreaterThan”。*/
function test() {
    let pvtFilter = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("score").PivotFilters.Item(1)
    console.log(pvtFilter.FilterType == xlCaptionIsGreaterThan)
}
```

```javascript
/*本示例显示活动工作表中第一个数据透视表的字段“name”的第一个筛选器应用的筛选器类型是否为“xlValueIsGreaterThan”，并通知用户。*/
function test() {
    let pvtFilter = ActiveSheet.PivotTables(1).PivotFields("name").PivotFilters.Item(1)
    if (pvtFilter.FilterType == xlValueIsGreaterThan) {
        console.log("筛选器类型是“xlValueIsGreaterThan”")
    } else {
        console.log("筛选器类型不是“xlValueIsGreaterThan”")
    }
}
```
