---
title: "AutoSortOrder 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/AutoSortOrder%20%E5%B1%9E%E6%80%A7.html
---
# AutoSortOrder (属性)

返回用于对指定数据透视表字段进行自动排序的次序。可为 **[XlSortOrder](./../../枚举/XlSortOrder 枚举.html)** 常量之一。**Long** 类型，只读。

## 示例

```javascript
/*本示例在消息框中显示“grades”字段的 AutoSort 参数值。*/
function test() {
    let pvtField = Worksheets.Item(1).PivotTables(1).PivotFields("grades")
    let aso = null
    switch (pvtField.AutoSortOrder) {
        case xlManual:
            aso = "manual"
            break;
        case xlAscending:
            aso = "ascending"
            break;
        case xlDescending:
            aso = "descending"
            break;
    }
    console.log(" sorted in " + aso + " order by " + pvtField.AutoSortField)
}
```

```javascript
/*本示例显示活动工作表中数据透视表是否按降序对字段“age”排序。*/
function test() {
    let pvtField = ActiveSheet.Range("I1").PivotTable.PivotFields("age")
    console.log(pvtField.AutoSortOrder == xlDescending)
}
```
