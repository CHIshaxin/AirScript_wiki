---
title: "Copy (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Copy%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BgYYE0MQ
---

# Copy (方法)

将单元格区域复制到指定的区域或剪贴板中。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Destination | any | 可选 | 指定区域要复制到的新域。如果省略此参数，ET 会将区域复制到剪贴板。 |

## 示例

```javascript
/*此示例将工作表 Sheet1 上单元格区域 A1:D4 中的公式复制到单元格区域 E5:H8 中。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:D4").Copy(Worksheets.Item("Sheet1").Range("E5"))
}
```

```javascript
/*此示例将显示工作表 Sheet1 上单元格 A1 中的数据类型并复制到单元格 B1 中。*/
function test() {
    Range("A1").Select()
    Range("B1").Copy(Selection)
}
```
