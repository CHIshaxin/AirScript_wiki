---
title: "DisplayFormat (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/DisplayFormat%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: B20JeZW_
---

# DisplayFormat (对象)

代表关联的 **[Range](./../Range/Range 对象.html)** 对象的显示设置。只读。

## 说明

更改区域的条件格式或表格样式等操作可导致当前用户界面中的显示内容与 **[Range](./../Range/Range 对象.html)** 对象相应属性中的值不一致。使用 **[DisplayFormat](./../DisplayFormat/DisplayFormat 对象.html)** 对象的属性可返回当前用户界面中显示的值。

## 示例

```javascript
/*本示例设置在当前用户界面中 G2 单元格边框的粗细为 xlThick ，并显示是否设置成功。*/
function test() {
    Range("G2").Style.Borders.Weight = xlThick
    console.log(Range("G2").DisplayFormat.Style.Borders.Weight == xlThick)
}
```

```javascript
/*本示例显示在当前用户界面中 D3:D5 单元格文本的字体颜色是否为红色。*/
function test() {
    console.log(Range("D3:D5").DisplayFormat.Font.Color == RGB(255, 0, 0))
}
```
