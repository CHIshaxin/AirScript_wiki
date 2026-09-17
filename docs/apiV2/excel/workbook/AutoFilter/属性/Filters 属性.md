---
title: "Filters 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AutoFilter/%E5%B1%9E%E6%80%A7/Filters%20%E5%B1%9E%E6%80%A7.html
---
# Filters (属性)

返回一个 **[Filters](./../../Filters/Filters 对象.html)** 集合，该集合表示自动筛选区域中的所有筛选器。只读。

## 示例

```javascript
/*下例将变量设置为工作表“Crew”上的筛选区域中第一列的筛选的 Criteria1 属性值。*/
function test()
{
  let mySheet = Application.Worksheets.Item("Crew")
  if(mySheet.AutoFilterMode){
      if(mySheet.AutoFilter.Filters.Item(1).On){
          let c1 = mySheet.AutoFilter.Filters.Item(1).Criteria1
      }
  }
}
```
