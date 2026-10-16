---
title: "Row (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Row%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DyB6WzJQ
---

# Row (属性)

返回区域中第一个子区域的第一行的行号。**Long** 类型，只读。

## 示例

```javascript
/*本示例将工作表 sheet1 中每隔一行的行高设置为 4 磅。*/
function test() {
    let rw = Application.Worksheets.Item("Sheet1").Rows
    for (let i = 1; i <= rw.Count; i++) {
        if ((rw.Item(i).Row) % 2 == 0) {
            rw.Item(i).RowHeight = 4
        }
    }
}
```

```javascript
/*本示例将 B1 单元格的值设置为 b1，并清除 A1 单元格所在行的内容*/
function test() {
    Range("B1").Value2 = 'b1'
    ActiveSheet.Rows.Item(Range("A1").Row).Clear()
}
```
