---
title: "Connector (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Connector%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D9TSK1lE
---

# Connector (属性)

如果指定的形状是连接符，则此属性为 **True**。**MsoTriState** 类型，只读。

## 示例

```javascript
/*本示例判断第一张工作表新建的shpRange对象是否是连接符。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([3])
    if (shpRange.Connector) {
        console.log("是连接符")
    } else {
        console.log("不是连接符")
    }
}
```

```javascript
/*本示例判断如果活动工作表新建的shpRange对象是连接符，则删除。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([3, 4])
    if (shpRange.Connector) {
        shpRange.Delete()
    }
}
```
