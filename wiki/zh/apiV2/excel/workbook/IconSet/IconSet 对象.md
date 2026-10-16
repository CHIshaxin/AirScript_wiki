---
title: "IconSet (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSet/IconSet%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CVm836nS
---

# IconSet (对象)

代表用于图标集条件格式规则的单一图标集。

## 说明

**[IconSet](./../IconSet/IconSet 对象.html)** 对象是 **[IconSets](./../IconSets/IconSets 对象.html)** 集合的子对象。

条件格式的图标集通过使用 **[IconSetCondition](./../IconSetCondition/IconSetCondition 对象.html)** 对象的 **IconSet** 属性来分配。通过以 **[Workbook](./../Workbook/Workbook 对象.html)** 对象的 **[IconSets](./../IconSets/IconSets 对象.html)** 属性的索引形式传递 **[XlIconSet](./../枚举/XlIconSet 枚举.html)** 枚举的其中一个常量，您可以将此属性设置为其中的一个内置图标集。有关详细信息，请参阅示例。

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

    // Change the icon set to a 5-arrow icon set
    iconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows)

    //The IconCriterion collection contains all of IconCriteria
    //By indexing into the collection you can modify each criteria
    let iconCriteria1 = iconSet.IconCriteria(1)
    iconCriteria1.Type = xlConditionValueNumber
    iconCriteria1.Value = 0
    iconCriteria1.Operator = 7

    let iconCriteria2 = iconSet.IconCriteria(2)
    iconCriteria2.Type = xlConditionValueNumber
    iconCriteria2.Value = 60
    iconCriteria2.Operator = 7

    let iconCriteria3 = iconSet.IconCriteria(3)
    iconCriteria3.Type = xlConditionValueNumber
    iconCriteria3.Value = 70
    iconCriteria3.Operator = 7

    let iconCriteria4 = iconSet.IconCriteria(4)
    iconCriteria4.Type = xlConditionValueNumber
    iconCriteria4.Value = 80
    iconCriteria4.Operator = 7

    let iconCriteria5 = iconSet.IconCriteria(5)
    iconCriteria5.Type = xlConditionValueNumber
    iconCriteria5.Value = 90
    iconCriteria5.Operator = 7

}
```

```javascript
/*本示例判断如果活动工作表上区域 B1:B10 中第一个图标集条件格式的图标集的名称是为三色旗，则显示该图标集中图标的数量。*/
function test() {
    let iconSet = ActiveSheet.Range("B1:B10").FormatConditions.Item(1).IconSet
    if (iconSet.ID == xl3Flags) {
        console.log(iconSet.Count)
    }
}
```
