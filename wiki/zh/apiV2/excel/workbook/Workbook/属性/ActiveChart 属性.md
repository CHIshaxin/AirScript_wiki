---
title: "ActiveChart (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/ActiveChart%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: _fe9SW7G
---

# ActiveChart (属性)

返回一个 **[Chart](./../../Chart/Chart 对象.html)** 对象，它代表活动图表（嵌入式图表或图表工作表）。嵌入式图表在被选中或激活时被认为是活动的。当没有图表处于活动状态时，此属性返回 **null**。

## 说明

如果不指定对象识别符，此属性返回活动工作簿中的活动图表。

## 示例

```javascript
/*本示例检查活动窗口中是否有图表被选中。*/
function test() {
    if (ActiveWindow.ActiveChart) {
        console.log(`活动图表名为 ${ActiveWindow.ActiveChart.Name}`)
    } else {
        console.log("活动窗口中未选中图表")
    }
}
```
