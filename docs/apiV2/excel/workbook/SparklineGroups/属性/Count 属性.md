---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroups/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回关联的 **[Range](./../../Range/Range 对象.html)** 对象中的迷你图组的计数。只读。

## 返回值

Long

## 示例

```javascript
/*本示例显示单元格区域 A1:A4 上迷你图组的数量。*/
function test() {
    console.log(Range("A1:A4").SparklineGroups.Count)
}
```
