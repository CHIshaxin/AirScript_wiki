---
title: "Characters (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E6%96%B9%E6%B3%95/Characters%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BCP-9yct
---

# Characters (方法)

返回 **[Characters](./../../Characters/Characters 对象.html)** 对象，它代表对象文本内某个区域的字符。使用 **[Characters](./../../Characters/Characters 对象.html)** 对象可为文本字符串内的字符设置格式。

## 说明

**[Characters](./../../Characters/Characters 对象.html)** 对象不是集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Start | any | 可选 | 要返回的第一个字符。如果此参数是 1 或被省略，则此属性返回一个以第一个字符为开头的字符区域。 |
| Length | any | 可选 | 要返回的字符数。如果省略此参数，则此属性返回字符串的后半部分（Start 字符之后的所有字符）。 |

## 示例

```javascript
/*本示例取消对图表工作表 Chart1 的标题加粗。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    charttitle.Characters().Font.Bold = false
}
```

```javascript
/*本示例显示图表工作表 Chart1 的标题的第三个字符。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    console.log(charttitle.Characters(2, 1).Text)
}
```
