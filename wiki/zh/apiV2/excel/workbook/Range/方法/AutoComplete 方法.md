---
title: "AutoComplete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/AutoComplete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BzT7xtaA
---

# AutoComplete (方法)

从列表中返回一个记忆式键入匹配项。如果没有相应的记忆式键入匹配项，或者在列表中有不止一个与已键入字符串相匹配的项，此方法将返回一空字符串。

## 说明

即使禁止了记忆式键入功能，该方法仍能工作。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| String | string | 必填 | 要完成的字符串。 |

## 返回值

String

## 示例

```javascript
/*本示例返回与“Ap”子串相匹配的记忆式键入匹配项。如果单元格 A5 所在列中包含一个与正在键入的字符串相匹配的数据项，将把该数据项作为记忆式键入匹配项返回。*/
function test() {
    let s = Worksheets.Item(1).Range("A5").AutoComplete("Ap")
    if (s.length > 0) {
        console.log("Completes to " + s)
    } else {
        console.log("Has no completion")
    }
}
```

```javascript
/*本示例将 A1 单元格的值设置为金山办公，显示 A6 单元格中与“金山”相匹配的数据项*/
function test() {
    Range("A1").Value2 = "金山办公"
    console.log(Range("A6").AutoComplete("金山"))
}
```
