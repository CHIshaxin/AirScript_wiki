---
title: "Comment 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/Comment%20%E5%B1%9E%E6%80%A7.html
---
# Comment (属性)

返回或设置与名称关联的批注。可读/写 **String** 类型。

## 说明

批注文本不能超过 255 个字符。

## 示例

```javascript
/*本示例在活动工作簿中为名称“test”添加备注。*/
function test() {
    ActiveWorkbook.Names.Item("test").Comment = "名称备注"
}
```

```javascript
/*本示例显示 Sheet1 中名称为“test1”的备注。*/
function test() {
    let name = Application.Worksheets.Item("Sheet1").Names.Item("test1")
    console.log(name.Comment)
}
```
