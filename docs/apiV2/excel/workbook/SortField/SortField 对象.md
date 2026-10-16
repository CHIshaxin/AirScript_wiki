---
title: "SortField 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/SortField%20%E5%AF%B9%E8%B1%A1.html
---
# SortField (对象)

**[SortField](./../SortField/SortField 对象.html)** 对象包含 **Worksheet**、**Lists** 和 **[AutoFilter](./../AutoFilter/AutoFilter 对象.html)** 对象的所有排序信息。

## 说明

开发人员可以使用 **BeforeSort** 事件重写 ET 的默认行为，将自己的排序算法写入应用程序中。

## 示例

```javascript
/*本示例遍历活动工作表的排序字段，将偶数位的排序字段设置为升序排序。*/
function test() {
    let sortfields = Application.ActiveSheet.Sort.SortFields
    for (let i = 1; i <= sortfields.Count; i++) {
        if (i % 2 == 0) {
            sortfields.Item(i).Order = xlDescending
        }
    }
}
```

```javascript
/*本示例在活动工作表新建排序字段，并显示该排序字段是否按字体颜色对数据进行升序排序。*/
function test() {
    let sortfield = Application.ActiveSheet.Sort.SortFields.Add(Range("A1:A3"), xlSortOnFontColor, xlAscending)
    if (sortfield.SortOn == xlSortOnFontColor && sortfield.Order == xlAscending) {
        console.log("排序字段按字体颜色对数据进行升序排序")
    } else {
        console.log("排序字段未按字体颜色和升序进行排序")
    }
}
```
