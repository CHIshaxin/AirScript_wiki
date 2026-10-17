---
title: "Worksheet (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/Worksheet%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CmWir2fP
---

# Worksheet (对象)

代表一个工作表。

## 说明

**[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象是 Worksheets 集合的成员。 **[Worksheets](./../Worksheets/Worksheets 对象.html)** 集合包含某个工作簿中所有的 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象。

**[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象也是 Sheets 集合的成员。 **[Sheets](./../Sheets/Sheets 对象.html)** 集合包含工作簿中所有的工作表（图表工作表和工作表）。

使用 Worksheets( *index*)（其中 *index* 是工作表索引号或名称）可返回一个 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象。

## 示例

```javascript
/*本示例隐藏活动工作簿中第一张工作表。*/
function test() {
    Application.Worksheets.Item(1).Visible = false
}
```

工作表索引号指示该工作表在工作簿的标签栏上的位置。`Worksheets(1)` 是工作簿中第一个（最左边的）工作表，而 `Worksheets(Worksheets.Count)` 是最后一个。所有工作表均包括在索引计数中，即便是隐藏工作表也是如此。

工作表名称显示在该工作表的标签上。使用 Name 属性可设置或返回工作表名称。

```javascript
/*本示例提醒用户输入密码，并用该密码保护工作表Sheet1上的方案。*/
function test() {
    let strPassword = "Enter the password for the worksheet"
    Application.Worksheets.Item("Sheet1").Protect(strPassword, null, null, true)
}
```

当工作表处于活动状态时，可以使用ActiveSheet属性来引用它。

```javascript
/*本示例激活工作表 Sheet1，将页面方向设置为横向，然后打印该工作表。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveSheet.PageSetup.Orientation = xlLandscape
    Application.ActiveSheet.PrintOut()
}
```
