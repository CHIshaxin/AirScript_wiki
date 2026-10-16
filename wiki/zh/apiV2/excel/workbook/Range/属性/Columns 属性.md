---
title: "Columns (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Columns%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DYAsNeWM
---

# Columns (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象代表指定区域中的列。

## 说明

在不使用对象识别符的情况下使用此属性等效于使用 ActiveSheet.Columns。

此属性在应用于一个是多重选定区域的 **[Range](./../../Range/Range 对象.html)** 对象时，会只从该区域的第一个子区域中返回列。例如，如果 **[Range](./../../Range/Range 对象.html)** 对象有两个子区域 A1:B2 和 C3:D4，那么，Selection.Columns.Count 的返回值是 2，而不是 4。若要对一个可能包含多重选定区域的区域使用此属性，请测试 Areas.Count 以确定此区域内是否包含多个子区域。如果包含，请对此区域内的每个子区域进行循环。

## 示例

```javascript
/*此示例将名为“myRange”区域第一列中每一单元格的值置为 0（零〕。*/
function test() {
    Application.Range("myRange").Columns.Item(1).Value2 = 0
}
```

```javascript
/*此示例显示 Sheet1 上选定区域中的列数。如果选择了多个子区域，此示例将对每一个子区域进行循环。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let areaCount = Application.Selection.Areas.Count
    console.log(areaCount)
    if (areaCount <= 1) {
        console.log("The selection contains " + Application.Selection.Columns.Count + " columns.")
    } else {
        for (let i = 1; i <= areaCount; i++) {
            console.log("Area " + i + " of the selection contains " + Application.Selection.Areas.Item(i).Columns.Count + " columns.")
        }
    }
}
```
