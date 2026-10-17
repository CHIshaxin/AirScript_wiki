---
title: "DefaultTableStyle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/DefaultTableStyle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DffYG7ij
---

# DefaultTableStyle (属性)

指定 **[TableStyles](./../../TableStyles/TableStyles 对象.html)** 集合中用作默认 TableStyle 的表样式。可读/写 **Variant** 类型。

## 示例

```javascript
/*本示例演示活动工作簿的默认样式*/
function test() {
    console.log(ActiveWorkbook.DefaultTableStyle)
}
```

```javascript
/*本示例演示向活动工作簿的 TableStyles 集合中添加 style1 样式，并指定 style1 为默认样式*/
function test() {
    ActiveWorkbook.TableStyles.Add("style1")
    ActiveWorkbook.DefaultTableStyle = "style1"
}
```
