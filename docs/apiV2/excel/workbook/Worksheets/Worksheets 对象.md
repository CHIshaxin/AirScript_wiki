---
title: "Worksheets 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheets/Worksheets%20%E5%AF%B9%E8%B1%A1.html
---
# Worksheets (对象)

指定的或活动工作簿中所有 Worksheet 对象的集合。每个 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象都代表一个工作表。

## 说明

**[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象也是 **[Sheets](./../Sheets/Sheets 对象.html)** 集合的成员。 **[Sheets](./../Sheets/Sheets 对象.html)** 集合包含工作簿中所有的工作表（图表工作表和工作表）。

使用 Worksheets 属性可返回 **[Worksheets](./../Worksheets/Worksheets 对象.html)** 集合。下例将所有工作表移到工作簿尾部。

```javascript
Application.Worksheets.Move(Application.Worksheets.Item(Application.Worksheets.Count))
```

使用 Add 方法可创建一个新工作表并将它添加到集合。下例将两个新工作表添加到活动工作簿的工作表一之前。

```javascript
Application.Worksheets.Add(Application.Worksheets.Item(1), undefined, 2)
```

使用 **Worksheets**(*index*)（其中 *index* 是工作表索引号或名称）可返回一个 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象。下例隐藏活动工作簿中的工作表一。

```javascript
Application.Worksheets.Item(1).Visible = false
```
