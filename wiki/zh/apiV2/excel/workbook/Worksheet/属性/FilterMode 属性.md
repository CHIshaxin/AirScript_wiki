---
title: "FilterMode (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/FilterMode%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: lpgLBD3E
---

# FilterMode (属性)

如果工作表处于筛选模式，则为 **True**。只读 **Boolean** 类型。

## 示例

```javascript
/*本示例显示活动工作表的筛选模式。*/
function test() {
    console.log(Application.ActiveSheet.FilterMode)
}
```

```javascript
/*本示例判断活动工作簿中每一张工作表是否处于筛选模式，并通知用户*/
function test() {
    for (let i = 1; i <= Worksheets.Count; i++) {
        let sheet = Application.Worksheets.Item(i)
        if (sheet.FilterMode) {
            console.log(`工作表 ${sheet.Name} 处于筛选模式`)
        } else {
            console.log(`工作表 ${sheet.Name} 未处于筛选模式`)
        }
    }
}
```
