---
title: "XlSlicerCrossFilterType 枚举"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/%E6%9E%9A%E4%B8%BE/XlSlicerCrossFilterType%20%E6%9E%9A%E4%B8%BE.html
---
# XlSlicerCrossFilterType 枚举

---

指定由指定的切片器缓存所使用的交叉筛选类型以及显示方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| **xlSlicerCrossFilterShowItemsWithDataAtTop** | 2 | 为此切片器缓存开启交叉筛选，对于连接到同一数据源的其他切片器中的筛选选择，没有数据的任何平铺都将变灰。此外，有数据的平铺将移到切片器的顶部。（默认） |
| **xlSlicerCrossFilterShowItemsWithNoData** | 3 | 为此切片器缓存开启交叉筛选，对于连接到同一数据源的其他切片器中的筛选选择，没有数据的任何平铺都将变灰。 |
| **xlSlicerNoCrossFilter** | 1 | 完全关闭交叉筛选，因此，无论其他切片器中的筛选选择如何，所有平铺都将显示，并处于活动状态（未变灰）。 |
