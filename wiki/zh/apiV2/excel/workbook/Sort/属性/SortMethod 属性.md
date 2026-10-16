---
title: "SortMethod (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sort/%E5%B1%9E%E6%80%A7/SortMethod%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: XQKL5sjw
---

# SortMethod (属性)

指定中文排序方法。可读/写 **[XlSortMethod](./../../枚举/XlSortMethod 枚举.html)** 类型。

## 说明

| XlSortMethod 可以是以下 **SortMethod** 常量之一: |
| --- |
| **xlStroke** |
| **xlPinYin** |

## 示例

```javascript
/*本示例设置第一张工作表按字符的汉语拼音顺序排序。*/
function test() {
    Application.Sheets.Item(1).Sort.SortMethod = xlPinYin
}
```

```javascript
/*本示例判断活动工作表排序时执行的排序类型，并通知用户。*/
function test() {
    if (Application.ActiveSheet.Sort.SortMethod == xlPinYin) {
        console.log("执行按字符的汉语拼音顺序排序")
    } else if (Application.ActiveSheet.Sort.SortMethod == xlStroke) {
        console.log("执行按每个字符的笔划数排序")
    }
}
```
