---
title: "Intersect (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/Intersect%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: MwEgbQEN
---

# Intersect (方法)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示两个或多个区域重叠的矩形区域。如果指定了不同工作表中的一个或多个区域，则返回错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Arg1 | Range | 必填 | 重叠的区域。必须指定至少两个 **[Range](./../../Range/Range 对象.html)** 对象。 |
| Arg2 | Range | 必填 | 重叠的区域。必须指定至少两个 **[Range](./../../Range/Range 对象.html)** 对象。 |
| Arg3 | any | 可选 | Arg30 - 相交范围。 |

## 返回值

Range

## 示例

```javascript
/*本示例在 Sheet1 中选定两个区域A1:D1和D1:F1的重叠区域。如果所选区域不重叠，本示例将显示一条相应的信息。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    let intersect = Application.Intersect(Range("A1:D1"), Range("D1:F1"))
    if (intersect === null) {
        console.log("Ranges do not intersect")
    } else {
        intersect.Select()
    }
}
```

```javascript
/*本示例比较 Worksheet.Range 属性、 Application.Union 方法和 Intersect 方法。*/
function test() {
    // Selects cells A1 to A10.
    Range("A1:A10").Select()
    // Selects cells A1 to A10.
    Range(Range("A1"), Range("A10")).Select()
    // Selects cells A1 and A10.
    Range("A1, A10").Select()
    // Selects cells A1 and A10.
    Union(Range("A1"), Range("A10")).Select()
    // Selects cell A5.
    Range("A1:A5 A5:A10").Select()
    // Selects cell A5.
    Intersect(Range("A1:A5"), Range("A5:A10")).Select()
}
```
