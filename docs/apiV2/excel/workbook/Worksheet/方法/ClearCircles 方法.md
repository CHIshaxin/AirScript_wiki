---
title: "ClearCircles 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/ClearCircles%20%E6%96%B9%E6%B3%95.html
---
# ClearCircles (方法)

清除指定工作表的无效数据项的圈释。

## 说明

可用 **CircleInvalid** 方法对含有无效数据的单元格进行圈释。

## 示例

```javascript
/*本示例清除对第一张工作表中无效数据项的圈释。*/
function test() {
    Worksheets.Item(1).ClearCircles()
}
```

```javascript
/*本示例清除打开的所有工作表中无效数据项的圈释。*/
function test() {
    for (let i = 1; i <= Worksheets.Count; i++) {
        Worksheets.Item(i).ClearCircles()
    }
}
```
