---
title: "AddCustomList (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/AddCustomList%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DNQzjmtl
---

# AddCustomList (方法)

为自定义自动填充和/或自定义排序添加自定义列表。

## 说明

如果要添加的列表已经存在，则本方法不起作用。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| ListArray | any | 必填 | 将源数据指定为字符串数组或 **[Range](./../../Range/Range 对象.html)** 对象。 |
| ByRow | any | 可选 | 仅当 *ListArray* 为 **[Range](./../../Range/Range 对象.html)** 对象时使用。如果为 **True**，则使用区域中的每一行创建自定义列表；如果为 **False**，则使用区域中的每一列创建自定义列表。如果省略该参数，并且区域中的行数比列数多（或者行数与列数相等），则 ET 使用区域中的每一列创建自定义列表。如果省略该参数，并且区域中的列数比行数多，则 ET 使用区域中的每一行创建自定义列表。 |

## 示例

```javascript
/*本示例将一个字符串数组添加到自定义列表中。*/
function test() {
    Application.AddCustomList(["cogs", "sprockets", "widgets", "gizmos"])
}
```

```javascript
/*本示例演示在A1:D1区域中分别赋值英文单词，并设置自定义列表。*/
function test() {
    Range("A1").Value2 = "apple"
    Range("B1").Value2 = "banana"
    Range("C1").Value2 = "watermelon"
    Range("D1").Value2 = "lemon"
    Application.AddCustomList(Range("A1:D1"), true)
}
```
