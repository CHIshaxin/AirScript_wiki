---
title: "Text (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/Text%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CW32nbLa
---

# Text (属性)

返回或设置一个 **String** 类型的值，该值指定条件格式规则所用的文本字符串。

## 说明

如果 **Type** 属性不是设置为 **xlTextString**，则忽略此属性。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个条件格式所用的特定文本。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(formatCondition.Text)
}
```

```javascript
/*本示例设置区域 A1:A10 中第一个条件格式所用的特定文本。*/
function test() {
    let formatCondition = Application.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.Text = "1"
}
```
