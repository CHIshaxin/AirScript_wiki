---
title: "On (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Filter/%E5%B1%9E%E6%80%A7/On%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: IRIZoUxv
---

# On (属性)

如果打开指定的筛选，则该属性值为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/* 下例将变量设置为工作表“Crew”上的筛选区域中第一列的筛选的 Criteria1 属性值*/
function test()
{
    let on1 = Application.Worksheets.Item("Crew")
    if(on1.AutoFilterMode) {
        let on2 = on1.AutoFilter.Filters.Item(1)
            if(on2.On) {
                let c1 = on2.Criteria1
            }
    }
}
```
