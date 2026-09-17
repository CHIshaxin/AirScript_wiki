---
title: "IconCriteria 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconCriteria/IconCriteria%20%E5%AF%B9%E8%B1%A1.html
---
# IconCriteria (对象)

代表 **IconCriterion** 对象的集合。每个 **IconCriterion** 代表图标集条件格式规则中每个图标的值和阈值类型。

## 说明

**[IconCriteria](./../IconCriteria/IconCriteria 对象.html)** 集合是从 **[IconSetCondition](./../IconSetCondition/IconSetCondition 对象.html)** 对象的 **IconCriteria** 属性中返回的。通过将索引传递到集合中，您可以访问集合中的每个 **IconCriterion** 对象。有关详细信息，请参阅示例。

## 示例

```javascript
/*本示例创建了一个代表测试分数的数字范围，然后对该范围应用了图标集条件格式规则。图标集的类型将从默认图标变为五箭头图标集。最后，将阈值类型从百分点修改为硬编码数字。*/
function test() {
    // Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 55
    ActiveSheet.Range("C2").Value2 = 92
    ActiveSheet.Range("C3").Value2 = 88
    ActiveSheet.Range("C4").Value2 = 77
    ActiveSheet.Range("C5").Value2 = 66
    ActiveSheet.Range("C6").Value2 = 93
    ActiveSheet.Range("C7").Value2 = 76
    ActiveSheet.Range("C8").Value2 = 80
    ActiveSheet.Range("C9").Value2 = 79
    ActiveSheet.Range("C10").Value2 = 83
    ActiveSheet.Range("C11").Value2 = 66
    ActiveSheet.Range("C12").Value2 = 74
    Range("C1:C12").Select()

    // Create an icon set conditional format for the created sample data range
    let iconSet = Selection.FormatConditions.AddIconSetCondition()

    // Change the icon set to a five-arrow icon set
    iconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows)

    // The IconCriterion collection contains all IconCriteria
    // By indexing into the collection you can modify each criterion
    let iconCriteria2 = iconSet.IconCriteria.Item(2)
    iconCriteria2.Type = xlConditionValueNumber
    iconCriteria2.Value = 60
    iconCriteria2.Operator = 7

    let iconCriteria3 = iconSet.IconCriteria.Item(3)
    iconCriteria3.Type = xlConditionValueNumber
    iconCriteria3.Value = 70
    iconCriteria3.Operator = 7

    let iconCriteria4 = iconSet.IconCriteria.Item(4)
    iconCriteria4.Type = xlConditionValueNumber
    iconCriteria4.Value = 80
    iconCriteria4.Operator = 7

    let iconCriteria5 = iconSet.IconCriteria.Item(5)
    iconCriteria5.Type = xlConditionValueNumber
    iconCriteria5.Value = 90
    iconCriteria5.Operator = 7
}
```

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个图标集条件格式的第二个IconCriterion的运算符是否为xlGreaterEqual。*/
function test() {
    let iconSet = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(iconSet.IconCriteria.Item(2).Operator == xlGreaterEqual)
}
```
