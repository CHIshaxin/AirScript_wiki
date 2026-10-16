---
title: "ReadingOrder (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/ReadingOrder%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D-ldXNTO
---

# ReadingOrder (属性)

返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。

## 示例

```javascript
/*本示例设置C2单元格的阅读次序为从右到左。*/
function test() {
    Range("C2").Style.ReadingOrder = xlRTL
}
```

```javascript
/*本示例显示活动工作簿第一个样式的阅读顺序是否从左到右。*/
function test() {
    console.log(ActiveWorkbook.Styles.Item(1).ReadingOrder == xlLTR)
}
```
