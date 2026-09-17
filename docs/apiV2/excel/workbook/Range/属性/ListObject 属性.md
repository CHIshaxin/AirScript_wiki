---
title: "ListObject 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/ListObject%20%E5%B1%9E%E6%80%A7.html
---
# ListObject (属性)

为 **[Range](./../../Range/Range 对象.html)** 对象返回一个 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象。 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象，只读。

## 示例

```javascript
/*本示例将 A1 单元格所在的表格区域调整为 A1:C5*/
function test() {
    Range("A1").ListObject.Resize(Range("A1:C5"))
}
```

```javascript
/*本示例将 A1 单元格所在表格删除*/
function test() {
    Range("A1").ListObject.Delete()
}
```
