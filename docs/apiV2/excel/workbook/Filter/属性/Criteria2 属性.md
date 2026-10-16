---
title: "Criteria2 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Filter/%E5%B1%9E%E6%80%A7/Criteria2%20%E5%B1%9E%E6%80%A7.html
---
# Criteria2 (属性)

返回已筛选区域中指定列的第二个筛选值。**Variant** 类型，只读。

## 说明

使用筛选的 **Criteria2** 属性时，如果该筛选没有使用两个条件，则将出现错误。在试图使用 **Operator** 属性之前，需保证 **[Filter](./../../Filter/Filter 对象.html)** 对象的 **Criteria2** 属性不为零 (0)。

## 示例

```javascript
/* 下面的示例将一个变量设置为 Criteria2 属性值。该属性是 Crew 工作表上已筛选区域中的第一列筛选的*/
function test()
{
    let c1 = Application.Worksheets.Item("Crew")
    if(c1.AutoFilterMode) {
        let c2 = c1.AutoFilter.Filters.Item(1)
            if(c2.On && c2.Operator) {
                let c3 = c2.Criteria2
            }
            else {
                let c3 = "Not set"
            }
    }
}
```
