---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElements/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CDKVbYps
---

# Count (属性)

返回集合中对象的数目。只读 **Long** 类型。

## 示例

```javascript
/*本示例显示第一张工作表上第一张列表的表样式元素的数目。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Count)
}
```
