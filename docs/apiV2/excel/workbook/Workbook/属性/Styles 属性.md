---
title: "Styles 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/Styles%20%E5%B1%9E%E6%80%A7.html
---
# Styles (属性)

返回一个 **[Styles](./../../Styles/Styles 对象.html)** 集合，该集合表示指定工作簿中的所有样式。只读。

## 示例

```javascript
/*本示例演示活动工作簿中样式的总数。*/
function test() {
    console.log("活动工作簿中样式的总数：" + ActiveWorkbook.Styles.Count)
}
```

```javascript
/*本示例演示活动工作簿中第三个样式的 Name 。*/
function test() {
    console.log(ActiveWorkbook.Styles.Item(3).Name)
}
```

```javascript
/*本示例从活动工作簿中删除用户定义的“Stock Quote Style”样式。*/
function test() {
    ActiveWorkbook.Styles.Item("Stock Quote Style").Delete()
}
```
