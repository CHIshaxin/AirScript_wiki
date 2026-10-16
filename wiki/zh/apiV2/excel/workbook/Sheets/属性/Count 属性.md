---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sheets/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: wIyPwqmO
---

# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*此示例显示活动工作簿中的工作表数量。*/
function test() {
    console.log(Application.ActiveWorkbook.Sheets.Count)
}
```

```javascript
/*此示例判断如果当前工作簿成功添加一个工作表，则显示结果。*/
function test() {
    let count = ThisWorkbook.Sheets.Count
    ThisWorkbook.Sheets.Add()
    if (ThisWorkbook.Sheets.Count == count + 1) {
        console.log("成功添加一个工作表")
    }
}
```
