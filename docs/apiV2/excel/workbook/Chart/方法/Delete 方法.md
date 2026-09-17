---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例遍历工作簿上所有图表工作表，判断如果标题为成绩表，则删除该图表工作表。*/
function test() {
    for (let i = 1; i <= Charts.Count; i++) {
        if (Charts.Item(i).HasTitle && Charts.Item(i).ChartTitle.Text == "成绩表") {
            Charts.Item(i).Delete()
        }
    }
}
```

```javascript
/*本示例删除图表工作表 Chart1。*/
function test() {
    Application.Charts.Item("Chart1").Delete()
}
```
