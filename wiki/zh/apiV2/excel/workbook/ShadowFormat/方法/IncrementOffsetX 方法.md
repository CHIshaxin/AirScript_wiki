---
title: "IncrementOffsetX (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E6%96%B9%E6%B3%95/IncrementOffsetX%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CBSTcXS7
---

# IncrementOffsetX (方法)

按指定磅值更改阴影的水平偏移量。可以使用 **OffsetX**属性设置阴影的绝对水平偏移量。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Increment | float | 必填 | 指定阴影的水平偏移量（以磅为单位）。正值表示阴影右移，负值表示阴影左移。 |

## 示例

```javascript
/*本示例使第一张工作表中形状三的阴影向左移动 3 磅。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    worksheet.Shapes.Item(3).Shadow.IncrementOffsetX(-3)
}
```

```javascript
/*本示例将活动工作表中第一个形状阴影向右移动 4 磅。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(1).Shadow
    shadow.IncrementOffsetX(4)
}
```
