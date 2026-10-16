---
title: "Apply (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/Apply%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: VN3vEd1E
---

# Apply (方法)

应用通过 **PickUp** 方法复制的指定形状格式。

## 示例

```javascript
/*此示例复制第一个工作表上第一个形状的格式，然后将复制的格式应用到第二个形状上。*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).PickUp()
    Application.Worksheets.Item(1).Shapes.Item(2).Apply()
}
```

```javascript
/*此示例复制活动工作表上第三个形状的格式，然后将复制的格式应用到第二个形状上。*/
function test() {
    ActiveSheet.Shapes.Item(3).PickUp()
    ActiveSheet.Shapes.Item(2).Apply()
}
```
