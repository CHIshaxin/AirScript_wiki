---
title: "Areas 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Areas%20%E5%B1%9E%E6%80%A7.html
---
# Areas (属性)

返回一个 **[Areas](./../../Areas/Areas 对象.html)** 集合，该集合表示多重区域选择中的所有区域。只读。

## 说明

对于单一选择区域，**Areas** 属性返回只包含一个对象的集合，即原始 **[Range](./../../Range/Range 对象.html)** 对象本身。对于多重选择区域，**Areas** 属性返回一个集合，该集合包含与每个选定区域相对应的对象。

## 示例

```javascript
/*本示例在用户选定多个区域并试图执行某一命令时显示提示信息。该示例必须在工作表上运行。*/
function test() {
    if (Application.Selection.Areas.Count > 1) {
        console.log("Cannot do this to a multi-area selection.")
    }
}
```

```javascript
/*此示例选中A1,C1,E1单元格，并展示选中了几个区域*/
function test() {
    Union(Range("A1"), Range("C1"), Range("E1")).Select()
    console.log(Application.Selection.Areas.Count)
}
```
