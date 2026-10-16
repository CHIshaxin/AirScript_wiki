---
title: "Font (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cf0aS2np
---

# Font (属性)

返回一个 **[Font](./../../Font/Font 对象.html)** 对象，该对象代表当前用户界面中显示的关联 **Range** 的字体。只读。

## 返回值

Font

## 示例

```javascript
/*此示例显示第一个工作表中 A3 单元格文本的字体大小。*/
function test() {
    console.log(Worksheets.Item(1).Range("A3").DisplayFormat.Font.Size)
}
```

```javascript
/*此示例显示 D4 单元格文本是否加粗。*/
function test() {
    console.log(Range("D4").DisplayFormat.Font.Bold)
}
```
