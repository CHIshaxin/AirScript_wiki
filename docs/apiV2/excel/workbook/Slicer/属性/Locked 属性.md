---
title: "Locked 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/Locked%20%E5%B1%9E%E6%80%A7.html
---
# Locked (属性)

返回或设置在包含指定切片器的工作表受到保护时，是否可以修改该切片器。可读/写。

## 说明

如果在工作表受到保护时切片器被锁定，则为 **True**；如果在工作表受到保护时可以修改切片器，则为 **False**。默认值为 **True**。

## 返回值

Boolean

## 示例

```javascript
/*本示例判断包含切片器的工作表 Sheet1 受到保护时，是否可以修改该工作表上第一张数据透视表的第二个切片器，并通知用户。*/
function test() {
    let slicer = Worksheets.Item("Sheet1").PivotTables(1).Slicers.Item(2)
    if (slicer.Locked) {
        console.log("不可以修改切片器")
    } else {
        console.log("可以修改切片器")
    }
}
```

```javascript
/*本示例为包含切片器的活动工作表在受到保护时，设置不可修改该工作表上第一张数据透视表的第二个切片器。*/
function test() {
    ActiveSheet.PivotTables(1).Slicers.Item(2).Locked = true
}
```
