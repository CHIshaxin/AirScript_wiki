---
title: "Apply (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sort/%E6%96%B9%E6%B3%95/Apply%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D85fjfW3
---

# Apply (方法)

根据当前应用的排序状态对区域进行排序。

## 示例

```javascript
/*此示例对第一张工作表中区域 A1:A5 的数据进行升序排序。*/
function test() {
    Application.Sheets.Item(1).Activate()
    let sort = Application.ActiveSheet.Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Range("A1:A5"))
    sort.SetRange(Range("A1:A5"))
    sort.Apply()
}
```

```javascript
/*此示例为活动工作表区域 B1:B3 赋值，并对区域 B1:B3 进行笔划数排序。*/
function test() {
    Range("B1").Value2 = "张三"
    Range("B2").Value2 = "李四"
    Range("B3").Value2 = "王五"
    let sort = Application.ActiveSheet.Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Range("B1:B3"))
    sort.SetRange(Range("B1:B3"))
    sort.SortMethod = xlStroke
    sort.Apply()
}
```
