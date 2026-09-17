---
title: "ModifySourceData 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sparkline/%E6%96%B9%E6%B3%95/ModifySourceData%20%E6%96%B9%E6%B3%95.html
---
# ModifySourceData (方法)

修改单个迷你图的源数据。

## 说明

单个迷你图的数据源区域必须是连续的。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Formula | string | 必填 | 包含源数据的区域。 |

## 返回值

void

## 示例

```javascript
/*本示例修改单元格 I1 中第一个迷你图组的第一个迷你图的源数据。*/
function test() {
    Range("I1").SparklineGroups(1).Item(1).ModifySourceData("Sheet1!A1:C1")
}
```

```javascript
/*本示例修改活动工作表上单元格 I1 中第一个迷你图组的第一个迷你图的源数据，并显示修改前后的源数据区域。*/
function test() {
    let sparkline = ActiveSheet.Range("I1").SparklineGroups(1).Item(1)
    console.log(`修改前的源数据区域：${sparkline.SourceData}`)
    sparkline.ModifySourceData("B1:D1")
    console.log(`修改后的源数据区域：${sparkline.SourceData}`)
}
```
