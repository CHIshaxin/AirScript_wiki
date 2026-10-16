---
title: "ModifyKey (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/%E6%96%B9%E6%B3%95/ModifyKey%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DzZA7hPz
---

# ModifyKey (方法)

修改字段中按其排序的键值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Key | Range | 必填 | 指定要修改的键。 |

## 示例

```javascript
/*本示例在第一张工作表创建新的排序字段，并将第一张工作表排序字段的键值地址修改为B2:B5。*/
function test() {
    Application.Sheets.Item(1).Activate()
    let sortfield = ActiveSheet.Sort.SortFields.Add(Range("B1:B10"), xlSortOnValues)
    sortfield.ModifyKey(Range("B2:B5"))
}
```

```javascript
/*本示例判断如果活动工作表 SortField 对象的数量大于1，则将活动工作表第二个排序字段的键值地址修改为C1:C10。*/
function test() {
    let sort = Application.ActiveSheet.Sort.SortFields
    if (sort.Count > 1) {
        sort.Item(2).ModifyKey(Range("C1:C10"))
    }
}
```
