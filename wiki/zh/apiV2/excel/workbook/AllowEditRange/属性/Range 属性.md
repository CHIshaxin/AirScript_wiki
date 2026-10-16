---
title: "Range (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AllowEditRange/%E5%B1%9E%E6%80%A7/Range%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CHPp3oKY
---

# Range (属性)

返回一个 Range代表，它代表在受保护工作表上可编辑的区域的子集。

## 示例

```javascript
/*本示例显示第一张工作表上第一个 AllowEditRange 对象的单元格数量。*/
function test() {
    console.log(Worksheets.Item(1).Protection.AllowEditRanges.Item(1).Range.Count)
}
```

```javascript
/*本示例选中活动工作表上第二个 AllowEditRange 对象的单元格。*/
function test() {
    ActiveSheet.Protection.AllowEditRanges.Item(2).Range.Select()
}
```
