---
title: "Icon (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Icon/Icon%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BDB1ji_z
---

# Icon (对象)

代表用于条件格式规则的图标集中的单个图标。

## 说明

**[Icon](./../Icon/Icon 对象.html)** 对象是从 **[IconSet](./../IconSet/IconSet 对象.html)** 对象的 **Item** 属性中返回的。

## 示例

```javascript
/*本示例显示第一张工作表上区域 B1:B10 中第一个图标集条件格式的图标集中所有图标的索引。*/
function test() {
    let iconSet = Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(1).IconSet
    for (let i = 1; i <= iconSet.Count; i++) {
        console.log(iconSet.Item(i).Index)
    }
}
```

```javascript
/*本示例显示第一张工作表上区域 B1:B10 中第二个图标集条件格式的图标集中第一个图标的实际类型。*/
function test() {
    let icon = Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(2).IconSet.Item(1)
    console.log(icon.constructor.name)
}
```
