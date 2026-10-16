---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SheetViews/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DM6CYqPY
---

# Count (属性)

返回集合中对象的数目。只读 **Long** 类型。

## 示例

```javascript
/*本示例显示活动窗口的工作表视图数量*/
function test() {
    let count = Application.ActiveWindow.SheetViews.Count
    console.log(`活动窗口的工作表视图数量是 ${count}`)
}
```
