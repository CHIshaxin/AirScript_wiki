---
title: "Names 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/Names%20%E5%B1%9E%E6%80%A7.html
---
# Names (属性)

返回一个 **[Names](./../../Names/Names 对象.html)** 集合，该集合代表活动工作簿中的所有名称。只读 **[Names](./../../Names/Names 对象.html)** 对象。

## 说明

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveWorkbook.Names`。

## 示例

```javascript
/*本示显示 Names 集合中对象的数量。*/
function test() {
    console.log("Names 集合中对象的数量为：" + Application.Names.Count)
}
```

```javascript
/*此示例为活动工作簿中 Sheet1 上的区域 A1:D3 定义一个新名称。*/
function test() {
    console.log("定义的新名称为：" +  Application.ActiveWorkbook.Names.Add("tempRange", "=Sheet1!$A$1:$D$3").Name)
}
```
