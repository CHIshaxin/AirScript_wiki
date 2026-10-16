---
title: "ConditionValue 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ConditionValue/ConditionValue%20%E5%AF%B9%E8%B1%A1.html
---
# ConditionValue (对象)

代表数据条条件格式规则计算最短数据条和最长数据条的方法。

## 说明

**[ConditionValue](./../ConditionValue/ConditionValue 对象.html)** 对象是使用 **[Databar](./../Databar/Databar 对象.html)** 对象的 **MaxPoint** 或 **MinPoint** 属性返回的。

通过使用 **Modify** 方法，您可以从默认设置（最低值表示最短的数据条，最高值表示最长的数据条）中更改计算类型。

## 示例

```javascript
/*本示例将创建一个数据范围，然后对该范围应用数据条。并使用ConditionValue对象将阈值的计算方式更改为百分点。*/
function test() {
    //Create a range of data with a couple of extreme values
    Application.ActiveSheet.Range("D1").Value2 = 1
    Application.ActiveSheet.Range("D2").Value2 = 45
    Application.ActiveSheet.Range("D3").Value2 = 50
    Application.ActiveSheet.Range("D2:D3").AutoFill(Range("D2:D8"))
    Application.ActiveSheet.Range("D9").Value2 = 500
    Range("D1:D9").Select()

    //Create a data bar with default behavior
    let databar = Application.Selection.FormatConditions.AddDatabar()
    console.log("Because of the extreme values, middle data bars are very similar")

    //The MinPoint and MaxPoint properties return a ConditionValue object
    //which you can use to change threshold parameters
    databar.MinPoint.Modify(xlConditionValuePercentile, 5)
    databar.MaxPoint.Modify(xlConditionValuePercentile, 75)

}
```

```javascript
/*本示例判断如果第一张工作表上区域 A1:A10 中第一个（数据条）条件格式的最短的数据条的类型为xlConditionValueNumber，则修改该数据条的计算方法。*/
function test() {
    let conditionValue = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1).MinPoint
    if (conditionValue.Type == xlConditionValueNumber) {
        conditionValue.Modify(xlConditionValuePercent, 30)
    }
}
```
