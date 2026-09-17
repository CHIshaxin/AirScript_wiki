---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFormulas/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示活动工作表中数据透视表的公式的数量。*/
function test() {
    let pvtFormulas = ActiveSheet.Range("I1").PivotTable.PivotFormulas
    console.log(`公式的数量：${pvtFormulas.Count}`)
}
```
