---
title: "InsertIndent (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/InsertIndent%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Dilw8pin
---

# InsertIndent (方法)

向指定的区域添加缩进量。

## 说明

如果用本方法将缩进量设置为一个小于 0（零）或大于 15 的值，将出错。

可用 **IndentLevel** 属性返回某区域的缩进量。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| InsertAmount | int4 | 必填 | 要添加到当前缩进量的缩进量。 |

## 示例

```javascript
/*本示例减少单元格 A10 的缩进量。*/
function test() {
    Range("A10").InsertIndent(-1)
}
```

```javascript
/*本示例将 C1:E5 单元格的缩进量增加 1，并显示 C1 单元格的缩进量*/
function test() {
    Range("C1:E5").InsertIndent(1)
    console.log(Range("C1").IndentLevel)
}
```
