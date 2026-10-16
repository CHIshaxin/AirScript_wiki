---
title: "IconSetCondition (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSetCondition/IconSetCondition%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BrcKOTOE
---

# IconSetCondition (对象)

代表图标集条件格式规则。

## 说明

所有条件格式对象均包含在 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合对象中，该集合对象是 **[Range](./../Range/Range 对象.html)** 集合的子项。您可以使用 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合的 **Add** 方法或 **AddIconSetCondition** 方法创建图标集格式规则。

每个图标集包含三个、四个或五个图标。您可以使用 **[Workbook](./../Workbook/Workbook 对象.html)** 对象的 **[IconSets](./../IconSets/IconSets 对象.html)** 属性返回 **[IconSets](./../IconSets/IconSets 对象.html)** 对象以指定其中一个内置图标集。然后按 **[IconCriteria](./../IconCriteria/IconCriteria 对象.html)** 对象的成员将图标集中每个单独的图标分配给范围中的值的子集。阈值的类型也是由此对象指定的。

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

    //The IconCriterion collection contains all IconCriteria
    //By indexing into the collection you can modify each criterion
    let iconCriterion1 = iconSet.IconCriteria(1)
    iconCriterion1.Type = xlConditionValueNumber
    iconCriterion1.Value = 0
    iconCriterion1.Operator = 7
    let iconCriterion2 = iconSet.IconCriteria(2)
    iconCriterion2.Type = xlConditionValueNumber
    iconCriterion2.Value = 60
    iconCriterion2.Operator = 7
    let iconCriterion3 = iconSet.IconCriteria(3)
    iconCriterion3.Type = xlConditionValueNumber
    iconCriterion3.Value = 70
    iconCriterion3.Operator = 7
    let iconCriterion4 = iconSet.IconCriteria(4)
    iconCriterion4.Type = xlConditionValueNumber
    iconCriterion4.Value = 80
    iconCriterion4.Operator = 7
    let iconCriterion5 = iconSet.IconCriteria(5)
    iconCriterion5.Type = xlConditionValueNumber
    iconCriterion5.Value = 90
    iconCriterion5.Operator = 7
}
```

```javascript
/*本示例判断如果活动工作表上单元格区域 A1:A10 的第二个图标集条件格式使用的图标集的名称为xl3Arrows，则修改该条件格式所应用于的单元格区域，并将该图标集条件格式设置为仅展示图标。*/
function test() {
    let iconSet = ActiveSheet.Range("A1:A10").FormatConditions.Item(2)
    if (iconSet.IconSet.ID == xl3Arrows) {
        iconSet.ModifyAppliesToRange(Range("A1:A4"))
        iconSet.ShowIconOnly = true
    }
}
```
