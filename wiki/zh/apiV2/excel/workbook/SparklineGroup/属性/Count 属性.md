---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CsNMob3K
---

# Count (属性)

返回迷你图组中的迷你图数。只读。

## 返回值

Long

## 示例

```javascript
/*本示例显示单元格 I1 上第一个迷你图组中迷你图的数量。*/
function test() {
    console.log(Range("I1").SparklineGroups.Item(1).Count)
}
```
