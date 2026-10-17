---
title: "Rows (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Rows%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: tqUfGejm
---

# Rows (属性)

返回一个 Range 对象，它代表指定工作表中的所有行。 **[Range](./../../Range/Range 对象.html)** 对象，只读。

## 说明

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveSheet.Rows`。

该属性在应用于是多个选定区域的 **[Range](./../../Range/Range 对象.html)** 对象时，只从该区域中第一个子区域内返回行。例如，如果 **[Range](./../../Range/Range 对象.html)** 对象有两个子区域：A1:B2 和 C3:D4，则 `Selection.Rows.Count` 返回 2 而不是 4。若要在一个可能包含多个选定区域的区域中使用此属性，请测试 `Areas.Count` 以确定该区域是否包含多个选择区域。如果是，请对此区域内的每个子区域进行循环，如第 2 个示例所示。

## 示例

```javascript
/*本示例删除 Sheet1 的第三行。*/
function test() {
    Application.Worksheets.Item("Sheet1").Rows.Item(3).Delete()
}
```

```javascript
/*本示例显示 Sheet1 选定区域中的行数。如果选择了多个子区域，本示例将对每一个子区域进行循环。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let areaCount = Application.Selection.Areas.Count
    if (areaCount <= 1) {
        console.log("The selection contains " + Application.Selection.Rows.Count + " rows.")
    } else {
        for (let i = 1; i <= Application.Selection.Areas.Count; i++) {
            console.log("Area " + i + " of the selection contains " + Application.Selection.Areas.Item(i).Rows.Count + " rows.")
        }
    }
}
```
