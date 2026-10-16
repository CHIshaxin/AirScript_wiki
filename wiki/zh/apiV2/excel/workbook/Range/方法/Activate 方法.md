---
title: "Activate (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Activate%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DHTZWibR
---

# Activate (方法)

激活单个单元格，该单元格必须处于当前选定区域内。要选择单元格区域，请使用 **Select** 方法。返回值为Variant类型。

## 示例

```javascript
/*本示例选定工作表 Sheet1 上的单元格区域 A1:C3，并激活单元格 B2。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Range("A1:C3").Select()
    Application.Range("B2").Activate()
}
```
