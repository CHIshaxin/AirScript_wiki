---
title: "Copy (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sheets/%E6%96%B9%E6%B3%95/Copy%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: C8gxw-AT
---

# Copy (方法)

将工作表复制到工作簿的另一位置。

## 说明

如果既不指定 *Before* 也不指定 *After*，则 ET 将新建一个工作簿，其中包含复制的工作表。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Before | any | 可选 | 将要在其之前放置所复制工作表的工作表。如果指定了 After，则不能指定 Before。 |
| After | any | 可选 | 将要在其之后放置所复制工作表的工作表。如果指定了 Before，则不能指定 After。 |

## 示例

```javascript
/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之前。*/
function test() {
    Application.Worksheets.Item("Sheet1").Copy(Application.Worksheets.Item("Sheet3"))
}
```

```javascript
/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之后。*/
function test() {
    Sheets.Item("Sheet1").Copy(null, Application.Sheets.Item("Sheet3"))
}
```
