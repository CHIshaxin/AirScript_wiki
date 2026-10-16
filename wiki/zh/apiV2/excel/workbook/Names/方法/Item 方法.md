---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Names/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BtyWVx9b
---

# Item (方法)

从 **[Names](./../../Names/Names 对象.html)** 集合返回一个 **[Name](./../../Name/Name 对象.html)** 对象。

## 说明

必须只指定这三个参数中的一个。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 要返回的定义名称的名称或编号。 |
| IndexLocal | any | 可选 | 以用户语言表示的定义名称的名称。如果使用该参数，将不转换名称。 |
| RefersTo | any | 可选 | 名称引用的内容。使用该参数以引用的内容标识名称。 |

## 示例

```javascript
/*此示例将活动工作簿中的 mySortRange 名称删除。*/
function test() {
    Application.ActiveWorkbook.Names.Item("mySortRange").Delete()
}
```

```javascript
/*本示例显示 Sheet1 中索引为1的名称是否可见。*/
function test() {
    console.log(Application.Worksheets.Item("Sheet1").Names.Item(1).Visible)
}
```
