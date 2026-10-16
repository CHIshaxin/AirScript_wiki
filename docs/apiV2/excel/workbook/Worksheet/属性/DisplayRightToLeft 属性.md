---
title: "DisplayRightToLeft 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/DisplayRightToLeft%20%E5%B1%9E%E6%80%A7.html
---
# DisplayRightToLeft (属性)

如果指定工作表是从右到左显示（而非从左到右），则为 **True**。如果对象从左到右显示，则为 **False**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例演示活动工作表是否为从右到左显示。*/
function test() {
    console.log(Application.ActiveSheet.DisplayRightToLeft)
}
```

```javascript
/*此示例判断活动工作簿每一张工作表是否从左到右显示，并通知用户。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        let sheet = Application.Sheets.Item(i)
        if (sheet.DisplayRightToLeft) {
            console.log(`工作表 ${sheet.Name} 为从右到左显示`)
        } else {
            console.log(`工作表 ${sheet.Name} 为从左到右显示`)
        }
    }
}
```
