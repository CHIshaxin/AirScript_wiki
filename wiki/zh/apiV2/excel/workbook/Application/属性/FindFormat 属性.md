---
title: "FindFormat (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/FindFormat%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DNC1B24C
---

# FindFormat (属性)

设置或返回要查找的单元格格式类型的搜索条件。

## 示例

```javascript
/*本示例将搜索条件设置为查找文字为 Arial、常规和 10 号字体的单元格，并通知用户。*/
function test() {
    // Establish search criteria.
    let font = Application.FindFormat.Font
    font.Name = "Arial"
    font.FontStyle = "Regular"
    font.Size = 10

    // Notify user.
    let font1 = Application.FindFormat.Font
    console.log(`${font1.Name} - ${font1.FontStyle} - ${font1.Size} font is what the search criteria is set to.`)
}
```

```javascript
/*本示例将搜索条件设置为内部颜色是红色的单元格。*/
function test() {
    Application.FindFormat.Interior.ColorIndex = 3
}
```
