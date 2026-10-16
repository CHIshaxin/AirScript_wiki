---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Charts/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例先创建空白图表工作表，然后显示活动工作簿中图表工作表的数量。*/
function test() {
    let chats = Application.Charts
    chats.Add()
    console.log(Application.Charts.Count)
}
```

```javascript
/*本示例显示活动工作簿中图表工作表的数量是否为 2 。*/
function test() {
    console.log(Application.Charts.Count == 2)
}
```
