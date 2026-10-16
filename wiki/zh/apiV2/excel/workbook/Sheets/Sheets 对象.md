---
title: "Sheets (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sheets/Sheets%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: kU964Ypi
---

# Sheets (对象)

指定的或活动工作簿中所有工作表的集合。

## 说明

**[Sheets](./../Sheets/Sheets 对象.html)** 集合可以包含 **Chart** 或 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象。

如果希望返回所有类型的工作表， **[Sheets](./../Sheets/Sheets 对象.html)** 集合就非常有用。如果仅需使用某一类型的工作表，请参阅该工作表类型的对象主题。

使用 **Sheets** 属性可返回 **[Sheets](./../Sheets/Sheets 对象.html)** 集合。

## 示例

```javascript
/*此示例打印当前活动工作簿上的所有工作表。*/
function test() {
    Application.Sheets.PrintOut()
}
```

使用 **Add**方法可创建一个新的工作表并将它添加到集合。

```javascript
/*此示例给活动工作簿添加两个图表工作表，将它们放在工作簿中的工作表二之后。*/
function test() {
    Application.Sheets.Add(null, Application.Sheets.Item(2), 2, xlChart)
}
```

使用 **Sheets**(*index*)（其中 *index* 是工作表名称或索引号）可返回一个 **Chart** 或 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象。

```javascript
/*此示例激活名为 Sheet1 的工作表。*/
function test() {
    Application.Sheets.Item("Sheet1").Activate()
}
```

使用 **Sheets**(*array*) 可指定多个工作表。下例将名为“Sheet4”和“Sheet5”的工作表移到工作簿的开头。

```javascript
/*此示例将名为 Sheet4 和 Sheet5 的工作表移动到工作簿的开头。*/
function test() {
    Application.Sheets.Item(["Sheet4", "Sheet5"]).Move(Application.Sheets.Item(1))
}
```
