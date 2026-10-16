---
title: "Merge (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Styles/%E6%96%B9%E6%B3%95/Merge%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BQMdiwlV
---

# Merge (方法)

将另一张工作簿中的样式合并到 **[Styles](./../../Styles/Styles 对象.html)** 集合中。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Workbook | any | 必填 | 一个 Workbook 对象，它代表包含待合并样式的工作簿。 |

## 返回值

Variant

## 示例

```javascript
/*此示例将工作簿 Template.xls 中的样式合并到活动工作簿中。*/
function test() {
    Application.ActiveWorkbook.Styles.Merge(Workbooks.Item("Template.xls"))
}
```

```javascript
/*此示例判断工作簿 Template2.xls 中的第一个样式字体颜色为蓝色，则将 Template2.xls 的样式合并到活动工作簿中。*/
function test() {
    if (Workbooks.Item("Template2.xls").Styles.Item(1).Font.Color == RGB(0, 0, 255)) {
        ActiveWorkbook.Styles.Merge(Workbooks.Item("Template2.xls"))
    }
}
```
