---
title: "ListHeaderRows (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/ListHeaderRows%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C9OChcgg
---

# ListHeaderRows (属性)

返回指定区域中标题行的数目。**Long** 类型，只读。

## 说明

在使用该属性前，先使用 **CurrentRegion** 属性查找区域的边界。

## 示例

```javascript
/*本示例将变量 rTbl 设置为活动单元格所在的当前区域，但不包含任何标题行。*/
function test() {
    let rTbl = ActiveCell.CurrentRegion
    // remove the headers from the range
    let iHdrRows = rTbl.ListHeaderRows
    if (iHdrRows > 0) {
        // resize the range minus n rows
        rTbl = rTbl.Resize(rTbl.Rows.Count - iHdrRows)
        // and then move the resized range down to
        // get to the first non-header row
        rTbl = rTbl.Offset(iHdrRows)
    }
}
```

```javascript
/*本示例显示 A1:D6 单元格有多少标题行，并把 A2:D2 单元格区域内容清除，再显示一次*/
function test() {
    console.log(Range("A1:D6").ListHeaderRows)
    Range("A2:D2").Clear()
    console.log(Range("A1:D6").ListHeaderRows)
}
```
