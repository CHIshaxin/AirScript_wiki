---
title: "AddDatabar 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E6%96%B9%E6%B3%95/AddDatabar%20%E6%96%B9%E6%B3%95.html
---
# AddDatabar (方法)

返回一个代表指定区域数据条条件格式规则的 **[Databar](./../../Databar/Databar 对象.html)** 对象。

## 说明

数据条可帮助您查看某个单元格相对于其他单元格的值。数据条的长度代表单元格中的值。较长的数据条代表较高的值，较短的数据条代表较低的值。数据条在辨认较高和较低数值，特别是有大量数据存在的情况下很有帮助，比如，辨认假日销售报告中的最高玩具销量和最低玩具销量。

## 示例

```javascript
/*本示例在活动工作表上单元格区域 A1:A10 上新增数据条条件格式。*/
function test() {
    ActiveSheet.Range("A1:A10").FormatConditions.AddDatabar()
}
```

```javascript
/*本示例选择一个单元格区域，将数据条条件格式规则添加到该区域，然后将数据条的填充色设置为纯色。*/
function test() {
    Range("B1:B10").Select()
    let databar = Selection.FormatConditions.AddDatabar()
    databar.BarFillType = xlDataBarFillSolid
}
```
