---
title: "TopBottom (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Top10/%E5%B1%9E%E6%80%A7/TopBottom%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CyewQ6b2
---

# TopBottom (属性)

返回或设置 **[XlTopBottom](./../../枚举/XlTopBottom 枚举.html)** 枚举的常量之一，该常量决定是从顶端还是从底端开始计算排位。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第二个（Top10）条件格式的TopBottom属性是否为xlTop10Top。*/
function test() {
    let top = ActiveSheet.Range("A1:A10").FormatConditions.Item(2)
    console.log(top.TopBottom == xlTop10Top)
}
```

```javascript
/*本示例设置第一张工作表上区域 A1:A10 中第一个（Top10）条件格式的TopBottom属性。*/
function test() {
    let top = Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1)
    top.TopBottom = xlTop10Bottom
}
```
