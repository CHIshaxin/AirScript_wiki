---
title: "Hyperlinks 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Hyperlinks%20%E5%B1%9E%E6%80%A7.html
---
# Hyperlinks (属性)

返回 **[Hyperlinks](./../../Hyperlinks/Hyperlinks 对象.html)** 集合，它代表区域的超链接。

## 示例

```javascript
/*本示例显示 A1 单元格的超链接数量*/
function test() {
    console.log(Range("A1").Hyperlinks.Count)
}
```

```javascript
/*本示例为 C2 单元格添加超链接 https://www.wps.cn/，并显示 C2 单元格的超链接数量*/
function test() {
    Range("C2").Hyperlinks.Add(Range("C2"), "https://www.wps.cn/")
    console.log(Range("C2").Hyperlinks.Count)
}
```
