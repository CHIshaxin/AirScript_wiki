---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: swhRZPPf
---

# Item (方法)

从集合中返回一个对象。

## 说明

对象的文本名称就是 **Name** 和 **Value** 属性的值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*此示例激活工作表 Sheet1 上第一张图表。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects().Item(1).Activate()
}
```

```javascript
/*此示例为第一张工作表上图表 Chart 2 添加圆角。*/
function test() {
    let chartobject = Application.Worksheets.Item(1).ChartObjects().Item("Chart 2")
    chartobject.RoundedCorners = true
}
```
