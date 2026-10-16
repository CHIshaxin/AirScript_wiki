---
title: "Sort 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/Sort%20%E5%B1%9E%E6%80%A7.html
---
# Sort (属性)

获取或设置 **[ListObject](./../../ListObject/ListObject 对象.html)** 集合的排序列和排序次序。

## 示例

```javascript
/*本示例显示活动工作表上包含 C2 单元格的列表的排序方向是否为 xlSortColumns。*/
function test() {
    let listObj = Application.ActiveSheet.Range("C2").ListObject
    console.log(listObj.Sort.Orientation == xlSortColumns)
}
```

```javascript
/*本示例将工作表 Sheet1 上第一张列表的排序方式设置为区分大小写的排序。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.Sort.MatchCase = true
}
```
