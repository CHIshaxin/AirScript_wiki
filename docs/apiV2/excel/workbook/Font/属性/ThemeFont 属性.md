---
title: "ThemeFont 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/ThemeFont%20%E5%B1%9E%E6%80%A7.html
---
# ThemeFont (属性)

返回或设置与指定对象关联的应用字体方案中的主题字体。可读/写 **[XlThemeFont](./../../枚举/XlThemeFont 枚举.html)** 类型。

## 示例

```javascript
/*本示例将Sheet1中活动单元格的字体的主题字体设置为主题正文*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveCell.Font.ThemeFont = xlThemeFontMinor
}
```

```javascript
/*本示例判断工作表Sheet1中单元格A5的字体的主题字体是否为主题标题，不是则设置其字体大小为16磅*/
function test() {
    let font = Application.Worksheets("Sheet1").Range("A5").Font
    if (font.ThemeFont == xlThemeFontMajor) {
        console.log("字体的主题字体是主题标题")
    } else {
        font.Size = 16
    }
}
```

\`\`
