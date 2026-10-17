---
title: "EnableSelection (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/EnableSelection%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BUUJGcRg
---

# EnableSelection (属性)

返回或设置工作表中可被选择的内容。 **[XlEnableSelection](./../../枚举/XlEnableSelection 枚举.html)** 类型，可读写。

## 说明

仅当工作表处于保护状态时，本属性才能起作用：**xlNoSelection** 防止在工作表上进行任何选择操作，**xlUnlockedCells** 仅允许 **Locked** 属性为 **False** 的单元格才能被选定，而 **xlNoRestrictions** 允许选定任意单元格。

## 示例

```javascript
/*本示例使在第一张工作表上不能进行任何选择。*/
function test() {
    Worksheets.Item(1).EnableSelection = xlNoSelection
    Worksheets.Item(1).Protect(undefined, undefined, true, undefined, true)
}
```

```javascript
/*本示例演示除A1：B1区域外，其他区域都可被选择。*/
function test() {
    Range("A1:B1").Locked = true
    ActiveSheet.EnableSelection = xlUnlockedCells
}
```
