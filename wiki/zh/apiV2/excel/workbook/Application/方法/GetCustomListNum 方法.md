---
title: "GetCustomListNum (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/GetCustomListNum%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: iIYa4TK-
---

# GetCustomListNum (方法)

返回字符串数组的自定义序列号。使用本方法既可对内置序列进行匹配，也可对自定义序列进行匹配。

## 说明

如果没有对应序列，则使用本方法将出错。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| ListArray | any | 必填 | 字符串数组。 |

## 返回值

Long

## 示例

```javascript
/*本示例删除一个自定义序列。*/
function test() {
    let customLN = Application.GetCustomListNum(["cogs", "sprockets", "widgets", "gizmos"])
    Application.DeleteCustomList(customLN)
}
```

```javascript
/*本示例显示自定义序列号。*/
function test() {
    console.log(Application.GetCustomListNum(["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]))
}
```
