---
title: "Hyperlinks 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlinks/Hyperlinks%20%E5%AF%B9%E8%B1%A1.html
---
# Hyperlinks (对象)

代表工作表或区域的超链接的集合。

## 说明

每个超链接都由一个 **[Hyperlink](./../Hyperlink/Hyperlink 对象.html)** 对象代表。

使用 **Hyperlinks** 属性可返回 **[Hyperlinks](./../Hyperlinks/Hyperlinks 对象.html)** 集合。

## 示例

```javascript
/*本示例检查并打开第一个工作表中包含"wps"单词所有的超链接。*/
function test() {
    let hyperlinks = Application.Worksheets.Item(1).Hyperlinks
    for (let i = 1; i <= hyperlinks.Count; i++) {
        if (hyperlinks.Item(i).Name.indexOf("wps") != -1) {
            hyperlinks.Item(i).Follow()
        }
    }
}
```

使用 **Add** 方法可创建一个超链接并将它添加到 **[Hyperlinks](./../Hyperlinks/Hyperlinks 对象.html)** 集合。

```javascript
/*本示例向第一张工作表的单元格E5中添加一个超链接。*/
function test() {
    let sheet = Application.Worksheets.Item(1)
    sheet.Hyperlinks.Add(sheet.Range("E5"), "https://kingsoft.com")
}
```
