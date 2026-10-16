---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AllowEditRanges/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
---
# Add (方法)

在受保护的工作表中添加可编辑的单元格区域。返回 AllowEditRange对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Title | string | 必填 | 单元格区域的标题。 |
| Range | Range | 必填 | Range 对象。允许编辑的单元格区域。 |
| Password | any | 可选 | 单元格区域的密码。 |

## 示例

```javascript
/*此示例在活动工作表上添加可编辑的单元格区域。*/
function test() {
    ActiveSheet.Protection.AllowEditRanges.Add("AllowEditRange3", Range("C1:C4"), "Y6dh!et5")
}
```

```javascript
/*本示例在第一张工作表上添加可编辑的单元格区域。*/
function test() {
    Worksheets.Item(1).Protection.AllowEditRanges.Add("AllowEditRange4", Range("D1:D4"))
}
```
