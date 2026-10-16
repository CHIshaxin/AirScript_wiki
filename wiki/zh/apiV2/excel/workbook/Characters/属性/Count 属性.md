---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Characters/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BdILeAge
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*此示例将 A1 单元格的最后一个字符设为上标字符。*/
function test(){
    let n = Application.Worksheets.Item("Sheet1").Range("A1").Characters.Count
    Application.Worksheets.Item("Sheet1").Range("A1").Characters(n, 1).Font.Superscript = true
}
```
