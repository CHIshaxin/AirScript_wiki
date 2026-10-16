---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: szy-DD6w
---

# Name (属性)

返回或设置一个 **Variant** 值，它代表对象的名称。

## 说明

**[Font](./../../Font/Font 对象.html)** 对象的名称为 **String**。

## 示例

```javascript
/*本示例将第一个工作表中单元格 B2:C3 区域的字体名称设置为黑体。*/
function test() {
    let font = Application.Worksheets.Item(1).Range("B2:C3").Font
    font.Name = "黑体"
}
```

```javascript
/*本示例判断第一个工作表中单元格E3的字体名称是否为宋体，是则将字形设置为倾斜，不是则设置字体名称为黑体*/
function test() {
    let font = Application.Worksheets.Item(1).Range("E3").Font
    if (font.Name == "宋体") {
        font.Italic = true
        console.log(`字体名称为宋体，已将字形设置为倾斜`)
    } else {
        font.Name = "黑体"
        console.log(`字体名称已设置为黑体`)
    }
}
```

``
