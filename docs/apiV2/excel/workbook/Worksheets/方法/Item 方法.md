---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheets/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 说明

有关返回集合中单个成员的详细信息，请参阅返回集合中的对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*此示例删除第一张工作表*/
function test() {
    Application.Worksheets.Item(1).Delete()
}
```

```javascript
/*本示例显示第三个工作表的名称。*/
function test() {
    console.log(Application.Worksheets.Item(3).Name)
}
```
