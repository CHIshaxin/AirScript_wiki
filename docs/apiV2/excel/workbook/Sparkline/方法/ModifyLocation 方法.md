---
title: "ModifyLocation 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sparkline/%E6%96%B9%E6%B3%95/ModifyLocation%20%E6%96%B9%E6%B3%95.html
---
# ModifyLocation (方法)

修改单个迷你图的位置。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Range | Range | 必填 | 包含迷你图的单元格。 |

## 返回值

void

## 示例

```javascript
/*本示例修改单元格 I1 中第一个迷你图组的第一个迷你图的位置。*/
function test() {
    Range("I1").SparklineGroups(1).Item(1).ModifyLocation(Range("F1"))
}
```

```javascript
/*本示例修改活动工作表上单元格 I1 中第一个迷你图组的第一个迷你图的位置。*/
function test() {
    let sparkline = ActiveSheet.Range("I1").SparklineGroups(1).Item(1)
    sparkline.ModifyLocation(Range("A10"))
}
```
