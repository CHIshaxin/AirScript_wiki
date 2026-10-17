---
title: "TableStyles (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/TableStyles%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CrE0m1_q
---

# TableStyles (属性)

返回当前工作簿的一个 **[TableStyles](./../../TableStyles/TableStyles 对象.html)** 集合对象，该对象引用当前工作簿中所使用的样式。只读。

## 说明

您可以处理由该属性返回的表样式。

## 示例

```javascript
/*本示例演示向活动工作簿中新加样式 style1 。*/
function test() {
    ActiveWorkbook.TableStyles.Add("style1")
}
```

```javascript
/*本示例演示活动工作簿中所使用的样式总数。*/
function test() {
    console.log(ActiveWorkbook.TableStyles.Count)
}
```
