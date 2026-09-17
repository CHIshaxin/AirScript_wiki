---
title: "Filters 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Filters/Filters%20%E5%AF%B9%E8%B1%A1.html
---
# Filters (对象)

由多个 **[Filter](./../Filter/Filter 对象.html)** 对象组成的集合，这些对象代表自动筛选区域内的所有筛选。

## 说明

使用 **Filters** 属性可返回 **[Filters](./../Filters/Filters 对象.html)** 集合。下例创建一个列表，其中包含工作表 Crew 中已自动筛选的区域内所有筛选的条件和运算符。

```javascript
function test()
{
    let f
    let op
    let c1
    let c2
    let ns = "Not set"

    let w = Application.Worksheets.Item("Sheet1")
    let w2 = Application.Worksheets.Item("Sheet2")
    let rw = 1
    for(let i = 1 ; i <= w.AutoFilter.Filters.Count; i++) {
        f = w.AutoFilter.Filters.Item(i)
        if(f.On) {
            let c1 = f.Criteria1.substr(f.Criteria1.length - 1,1)
            if(f.Operator) {
                op = f.Operator
                c2 = f.Criteria2.substr(f.Criteria2.length - 1,1)
            } else {
                op = ns
                c2 = ns
            }
        } else {
            c1 = ns
            op = ns
            c2 = ns
        }
        w2.Cells.Item(rw, 1).Value2 = c1
        w2.Cells.Item(rw, 2).Value2 = op
        w2.Cells.Item(rw, 3).Value2 = c2
        rw = rw + 1
    }
}
```

使用 **Filters**(*index*)（其中 *index* 是筛选的标题或索引号）可返回单个 **[Filter](./../Filter/Filter 对象.html)** 对象。下例将一个变量设置为工作表 Crew 上的筛选区域中第一列的筛选的 **On** 属性的值。

```javascript
function test()
{
    let w = Worksheets.Item("Crew")
    if(w.AutoFilterMode) {
        let filterIsOn = w.AutoFilter.Filters.Item(1).On
    }
}
```
