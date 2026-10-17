---
title: "XlEnableCancelKey 枚举"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/%E6%9E%9A%E4%B8%BE/XlEnableCancelKey%20%E6%9E%9A%E4%B8%BE.html
updated: 1789700916000
hash: CGiBArYl
---

# XlEnableCancelKey 枚举

---

指定 WPS Office ET 2007 如何处理 Ctrl+Break（或 Esc、Command+Period）用户中断以用于运行程序。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| **xlDisabled** | 0 | 完全禁用“取消”键捕获功能。 |
| **xlErrorHandler** | 2 | 将中断作为错误发送给运行程序，由 On Error GoTo 语句设置的错误处理程序捕获。可捕获的错误代码为 18。 |
| **xlInterrupt** | 1 | 中断当前运行程序，用户可进行调试或结束程序的运行。 |
