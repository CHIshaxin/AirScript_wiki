---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: rew64lkb
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

包含在集合中的一个 **[Shape](./../../Shape/Shape 对象.html)** 对象。

## 示例

```javascript
/*本示例将第一张工作表中新建的ShapeRange对象中第二个形状的宽度减小一半。如果变量 shpRange 不代表 ShapeRange 对象，则本示例无效。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2, 3])
    shpRange.Item(2).Width *= 0.5
}
```

```javascript
/*本示例将活动工作表中新建的ShapeRange对象中的第一个形状删除。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 3])
    shpRange.Item(1).Delete()
}
```
