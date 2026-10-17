---
title: "XlRobustConnect 枚举"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/%E6%9E%9A%E4%B8%BE/XlRobustConnect%20%E6%9E%9A%E4%B8%BE.html
updated: 1789700916000
hash: BKjlmBym
---

# XlRobustConnect 枚举

---

指定数据透视表缓存与其数据源连接的方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| **xlAlways** | 1 | 缓存始终使用外部源信息（由 **SourceConnectionFile** 或 **SourceDataFile** 属性定义）进行重新连接。 |
| **xlAsRequired** | 0 | 缓存通过 **Connection** 属性使用外部源信息进行重新连接。 |
| **xlNever** | 2 | 缓存从不使用源信息进行重新连接。 |
