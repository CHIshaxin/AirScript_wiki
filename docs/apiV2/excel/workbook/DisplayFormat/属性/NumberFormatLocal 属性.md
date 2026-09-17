---
title: "NumberFormatLocal 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/NumberFormatLocal%20%E5%B1%9E%E6%80%A7.html
---
# NumberFormatLocal (属性)

返回一个值，该值表示当前用户界面中以用户语言将关联的 **Range** 显示为字符串时的格式代码。只读。

## 返回值

Variant

## 示例

```javascript
/*本示例显示 D3 单元格用户界面中展示的用户语言的格式代码。*/
function test() {
    console.log(Range("D3").DisplayFormat.NumberFormatLocal)
}
```

```javascript
/*本示例显示 D3 单元格用户界面中展示的用户语言的格式代码是否为0.00，并通知用户。*/
function test() {
    if (Range("D5:D10").DisplayFormat.NumberFormatLocal == "0.00") {
        console.log("用户语言的格式代码是0.00")
    } else {
        console.log("用户语言的格式代码不是0.00")
    }
}
```
