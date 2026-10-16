---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortFields/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BStazj4R
---

# Item (方法)

返回一个 **[SortField](./../../SortField/SortField 对象.html)** 对象，该对象代表可在工作簿中排序的项的集合。只读。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 排序字段的索引值。 |

## 示例

```javascript
/*本示例判断如果活动工作表SortField对象数量大于1，则删除第二个SortField对象。*/
function test() {
    let sortfields = Application.ActiveSheet.Sort.SortFields
    if (sortfields.Count > 1) {
        sortfields.Item(2).Delete()
    }
}
```

```javascript
/*本示例在活动工作表新建排序字段，并显示第一个SortField对象的键值地址。*/
function test() {
    let sortfields = Application.ActiveSheet.Sort.SortFields
    sortfields.Add(Range("A1:A4"), xlSortOnValues)
    console.log(sortfields.Item(1).Key.Address())
}
```
