---
title: "DeleteCustomList (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/DeleteCustomList%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: TWyST2t6
---

# DeleteCustomList (方法)

删除一个自定义序列。

## 说明

如果列表数字小于 5 或者没有匹配的自定义序列，则使用本方法将产生错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| ListNum | int4 | 必填 | 自定义序列数字。此数字必须大于或等于 5（ET 有四个不可删除的内置自定义序列）。 |

## 示例

```javascript
/*本示例删除一个自定义序列。*/
function test() {
    let n = Application.GetCustomListNum(["water", "lemon", "apple"])
    Application.DeleteCustomList(n)
}
```

```javascript
/*本示例演示添加一个自定义序列，通知用户后，删除此自定义序列。*/
function test() {
    let customlist = ["water", "lemon", "apple"]
    Application.AddCustomList(customlist)
    console.log("添加成功，即将删除此自定义序列！")
    let n = Application.GetCustomListNum(customlist)
    Application.DeleteCustomList(n)
}
```
