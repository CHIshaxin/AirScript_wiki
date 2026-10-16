---
title: "ID 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSet/%E5%B1%9E%E6%80%A7/ID%20%E5%B1%9E%E6%80%A7.html
---
# ID (属性)

返回 **[XlIconSet](./../../枚举/XlIconSet 枚举.html)** 枚举的常量之一，该常量指定用于图标集条件格式规则的图标集的名称。

## 示例

```javascript
/*本示例显示第一张工作表上区域 B1:B10 中第二个图标集条件格式的图标集的名称是否为三向箭头。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(2).IconSet
    console.log(iconSet.ID == xl3Arrows)
}
```

```javascript
/*本示例判断活动工作表上区域 B1:B10 中第一个图标集条件格式的图标集的名称是否为三色旗，并通知用户。*/
function test() {
    let iconSet = ActiveSheet.Range("B1:B10").FormatConditions.Item(1).IconSet
    if (iconSet.ID == xl3Flags) {
        console.log("图标集的名称为三色旗")
    } else {
        console.log("图标集的名称不为三色旗")
    }
}
```
