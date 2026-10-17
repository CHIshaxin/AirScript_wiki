---
title: "CodeName (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/CodeName%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BoAZ1jxL
---

# CodeName (属性)

返回对象的代码名。**String** 型，只读。

## 说明

在“属性”窗口中“（名称）”右边的单元格中显示的值是所选对象的代码名。可以在设计过程中通过更改该值来改变对象的代码名。不能在运行过程中更改该属性。

对于一个返回指定对象的表达式，该表达式可使用对象的代码名。例如，如果第一张工作表的代码名为 Sheet1，则下列表达式是等价的。

```javascript
Worksheets(1).Range("a1")
Sheet1.Range("a1")
```

工作表的名称可以与其代码名不同。创建一张工作表时，其工作表名称和代码名是相同的，不过，更改工作表名称时并不影响其代码名，并且，更改工作表代码名（在宏编辑器中使用“属性”窗口）也不影响其名称。

## 示例

```javascript
/*此示例显示第一张工作表的代码名。*/
function test() {
    console.log(Worksheets.Item(1).CodeName)
}
```

```javascript
/*此示例演示活动工作表代码名的长度。*/
function test() {
    console.log(ActiveSheet.CodeName.length)
}
```
