---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStops/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回指定对象的计数。只读。

## 返回值

Long

## 示例

```javascript
/*以下示例显示活动单元格中的 ColorStops 数。*/
function test() {
    console.log(ActiveCell.Interior.Gradient.ColorStops.Count)
}
```
