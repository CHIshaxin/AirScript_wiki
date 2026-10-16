---
title: "SortOnValue 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/%E5%B1%9E%E6%80%A7/SortOnValue%20%E5%B1%9E%E6%80%A7.html
---
# SortOnValue (属性)

返回指定 **[SortField](./../../SortField/SortField 对象.html)** 对象执行排序的值。只读。

## 说明

**SortOnValue** 可与值、单元格颜色、字体颜色和单元格图标一起使用。

## 示例

```javascript
/*本示例按字体颜色对 Sheet1 上的 B 列中的数据进行升序排序。*/
function test() {
    Application.Worksheets.Item("Sheet1").Sort.SortFields.Clear()
    Application.Worksheets.Item("Sheet1").Sort.SortFields.Add(Range("B1:B25"), xlSortOnFontColor, xlAscending, null, xlSortNormal).SortOnValue.Color = RGB(0, 0, 0)

    let sort = Application.Worksheets.Item("Sheet1").Sort
    sort.SetRange(Range("A1:B25"))
    sort.Header = xlGuess
    sort.MatchCase = false
    sort.Orientation = xlTopToBottom
    sort.SortMethod = xlPinYin
    sort.Apply()
}
```

```javascript
/*本示例在活动工作表新建排序字段，并设置该字段分别按照单元格颜色等方式对数据进行排序。*/
function test() {
    let sortfield = Application.ActiveSheet.Sort.SortFields.Add(Range("C1:C5"), xlSortOnCellColor)
    sortfield.SortOnValue.Color = RGB(255, 255, 0)

    //设置该字段按字体颜色对数据进行排序
    sortfield.Sorton = xlSortOnFontColor
    sortfield.SortOnValue.Color = RGB(255, 255, 0)

    //设置该字段按图标对数据进行排序
    sortfield.Sorton = xlSortOnIcon
    sortfield.SortOnValue.Color = RGB(255, 255, 0)
    sortfield.SetIcon(ActiveWorkbook.IconSets(1).Item(3))
}
```
