---
title: "GetCustomListContents (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/GetCustomListContents%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B-5Rq2fI
---

# GetCustomListContents (方法)

返回一个自定义序列（一个字符串数组）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| ListNum | int4 | 必填 | 列表数字。 |

## 返回值

Variant

## 示例

```javascript
/*本示例将第一个自定义序列的元素写到工作表 Sheet1 的第一列中。*/
function test() {
    let listArray = Application.GetCustomListContents(1)
    for (let i = 1; i <= listArray.length; i++) {
        Worksheets.Item("Sheet1").Cells(i, 1).Value2 = listArray[i - 1]
    }
}
```

```javascript
/*本示例显示第二个自定义序列的元素。*/
function test() {
    console.log(Application.GetCustomListContents(2).toString())
}
```
