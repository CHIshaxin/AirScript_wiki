---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sheets/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*此示例激活工作表 Sheet1。*/
function test() {
    Application.Sheets.Item("Sheet1").Activate()
}
```

```javascript
/*本示例显示活动工作簿中的每个工作表的名称。*/
function test() {
    let sheets = Application.ActiveWorkbook.Sheets
    for (let i = 1; i <= sheets.Count; i++) {
        console.log(sheets.Item(i).Name)
    }
}
```
