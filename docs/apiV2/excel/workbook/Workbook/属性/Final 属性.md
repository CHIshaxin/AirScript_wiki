---
title: "Final 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/Final%20%E5%B1%9E%E6%80%A7.html
---
# Final (属性)

返回或设置一个 **Boolean** 类型的值，该值指示工作簿是否是最终的。可读/写 **Boolean** 类型。

## 说明

将值设置为 **True** 可将工作簿标记为最终版本，通知收件人（如果有）工作簿是最终的，并将工作簿设置为只读。

## 示例

```javascript
/*本示例演示将活动工作簿标记为最终版本。*/
function test() {
    ActiveWorkbook.Final = true
}
```

```javascript
/*本示例演示第一个工作簿是否为最终版本，以及Final属性的数据类型。*/
function test() {
    console.log(`当前工作簿是否为最终版本：${Workbooks.Item(1).Final}\nFinal属性的数据类型为：${typeof Workbooks.Item(1).Final}`)
}
```
