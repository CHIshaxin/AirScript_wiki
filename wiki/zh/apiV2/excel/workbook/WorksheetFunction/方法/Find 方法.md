---
title: "Find (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Find%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BKutpbrS
---

# Find (方法)

在工作表中查找特定信息。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | string | 必填 | 工作表名称。 |
| Arg2 | string | 必填 | 区域名称。 |
| Arg3 | any | 可选 | 精确限制查询的参数名称。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值B1等单元格，使用 Find 方法查找参数1字符串开始的位置。*/
function test() {
    Range("B1").Value2 = "World"
    Range("B2").Value2 = "hello,World!"
    Range("B3").Value2 = 4
    console.log(WorksheetFunction.Find(Range("B1").Value2, Range("B2").Value2, Range("B3").Value2))
}
```

```javascript
/*本示例使用 Find 方法查找参数1字符串开始的位置，并分别赋值给E1等单元格。*/
function test() {
    Range("E1").Value2 = WorksheetFunction.Find("in", "Believe in yourself.")
    Range("E2").Value2 = WorksheetFunction.Find("bird", "The bird who dares to fall is bird who learns to fly.", 2)
}
```
