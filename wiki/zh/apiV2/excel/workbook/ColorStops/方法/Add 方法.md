---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStops/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DjdRlVqC
---

# Add (方法)

将一个 **[ColorStop](./../../ColorStop/ColorStop 对象.html)** 对象添加到指定集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Position | double | 必填 | 代表要应用 **[ColorStop](./../../ColorStop/ColorStop 对象.html)** 的位置。 |

## 返回值

ColorStop

## 示例

```javascript
/*为当前选定区域添加一个ColorStop。*/
function test() {
    Range("A1:A10").Select()
    let colorStop = Selection.Interior.Gradient.ColorStops.Add(1)
    colorStop.ThemeColor = xlThemeColorAccent1
    colorStop.TintAndShade = 0
}
```

```javascript
/*本示例为活动单元格添加一个ColorStop，并将主题颜色设置为xlThemeColorDark2。*/
function test() {
    let colorStop = ActiveCell.Interior.Gradient.ColorStops.Add(0.5)
    colorStop.ThemeColor = xlThemeColorDark2
}
```
