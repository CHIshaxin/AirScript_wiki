---
title: "Criteria1 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Filter/%E5%B1%9E%E6%80%A7/Criteria1%20%E5%B1%9E%E6%80%A7.html
---
# Criteria1 (属性)

返回筛选区域内指定列上的第一个筛选值。**Variant** 类型，只读。

## 示例

```javascript
/* 下例将变量设置为工作表“Crew”上的筛选区域中第一列的筛选的 Criteria1 属性值*/
function test()
{
    let c1 = Application.Worksheets.Item("Crew")
    if(c1.AutoFilterMode) {
        let c2 = c1.AutoFilter.Filters.Item(1)
            if(c2.On) {
                c1 = c2.Criteria1
            }
    }
}
```
