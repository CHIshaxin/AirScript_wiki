---
title: "DeleteNumberFormat 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E6%96%B9%E6%B3%95/DeleteNumberFormat%20%E6%96%B9%E6%B3%95.html
---
# DeleteNumberFormat (方法)

从工作簿中删除一个自定义数字格式。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| NumberFormat | string | 必填 | 命名要删除的数字格式。 |

## 示例

```javascript
/*本示例从当前工作簿中删除“000-00-0000”数字格式。*/
function test() {
    ActiveWorkbook.DeleteNumberFormat("000-00-0000")
}
```

```javascript
/*本示例演示向Sheet1的A17单元格添加 hh:mm:ss 数字格式，告知用户后删除 hh:mm:ss 数字格式。*/
function test() {
    Worksheets.Item("Sheet1").Range("A17").NumberFormat = "hh:mm:ss"
    console.log("已添加hh:mm:ss数字格式")
    ActiveWorkbook.DeleteNumberFormat("hh:mm:ss")
}
```
