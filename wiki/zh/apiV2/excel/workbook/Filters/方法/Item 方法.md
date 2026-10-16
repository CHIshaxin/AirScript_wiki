---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Filters/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cqkw6DgR
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | int4 | 必填 | 对象的索引号。 |

## 示例

```javascript
/* 下例将一个变量设置为工作表 Crew 上的筛选区域中第一列的筛选的 On 属性的值*/
function test()
{
    let w = Worksheets.Item("Crew")
    if(w.AutoFilterMode) {
        filterIsOn = w.AutoFilter.Filters.Item(1).On
    }
}
```
