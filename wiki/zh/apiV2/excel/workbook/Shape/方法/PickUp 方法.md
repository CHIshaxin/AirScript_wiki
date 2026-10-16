---
title: "PickUp (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/PickUp%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CdN18ypT
---

# PickUp (方法)

复制指定形状的格式。使用 **Apply** 方法可将复制的格式应用到其他形状。

## 示例

```javascript
/*本示例将活动工作表第一个形状的格式应用于其它形状并删除第一个形状。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.Item(1).PickUp()
    shapes.Item(2).Apply()
    shapes.Item(3).Apply()
    shapes.Item(1).Delete()
}
```

```javascript
/*本示例首先复制第一张工作表上第一个形状的格式，然后将复制的形状格式应用于第二个形状。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Item(1).PickUp()
    shapes.Item(2).Apply()
}
```
