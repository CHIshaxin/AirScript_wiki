---
title: "ScrollArea (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/ScrollArea%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DRkJMd3P
---

# ScrollArea (属性)

以 A1 样式的区域引用形式返回或设置允许滚动的区域。用户不能选定滚动区域之外的单元格。**String** 类型，可读写。

## 说明

可将该属性设置为空字符串 ("") 以允许对整张工作表内所有单元格的选定。

## 示例

```javascript
/*本示例设置第一张工作表的滚动区域。*/
function test() {
    Worksheets.Item(1).ScrollArea = "A1:F10"
}
```

```javascript
/*本示例设置允许用户对活动工作簿中每一张工作表内所有单元格的选定，并通知用户。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        let sheet = Application.Worksheets.Item(i)
        sheet.ScrollArea = ""
        console.log(`允许对 ${sheet.Name} 整张工作表内所有单元格的选定`)
    }
}
```
