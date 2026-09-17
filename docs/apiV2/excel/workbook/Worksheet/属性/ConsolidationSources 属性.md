---
title: "ConsolidationSources 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/ConsolidationSources%20%E5%B1%9E%E6%80%A7.html
---
# ConsolidationSources (属性)

返回一个字符串数组，这些字符串是工作表中当前合并计算的数据源的名称。如果工作表中没有合并计算，将返回 **undefined**。**Variant** 类型，只读。

## 示例

```javascript
/*本示例显示在工作表 Sheet1 中用于合并计算的数据源的名称。该列表将出现在本示例新建的工作表中。*/
function test() {
    let newSheet = Worksheets.Add()
    newSheet.Range("A1").Value2 = "Consolidation Sources"
    let aSources = Worksheets.Item("Sheet1").ConsolidationSources
    if (aSources == null) {
        newSheet.Range("A2").Value2 = "none"
    } else {
        for (let i = 1; i <= aSources.length; i++) {
            newSheet.Cells(i + 1, 1).Value2 = aSources[i]
        }
    }
    newSheet.Columns("A:B").AutoFit()
}
```

```javascript
/*本示例判断活动工作表如果存在合并计算的数据源的名称，则输出ConsolidationSources属性的数据类型。*/
function test() {
    if (ActiveSheet.ConsolidationSources != null) {
        console.log("ConsolidationSources属性数据类型为：" + typeof ActiveSheet.ConsolidationSources)
    }
}
```
