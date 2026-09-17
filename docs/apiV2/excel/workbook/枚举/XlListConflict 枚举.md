---
title: "XlListConflict 枚举"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/%E6%9E%9A%E4%B8%BE/XlListConflict%20%E6%9E%9A%E4%B8%BE.html
---
# XlListConflict 枚举

---

指定冲突（用 ET 工作表中的列表的更改更新 Microsoft SharePoint Foundation 网站上的列表时）解决方法选项。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| **xlListConflictDialog** | 0 | 显示一个对话框，允许用户选择解决冲突的方式。 |
| **xlListConflictDiscardAllConflicts** | 2 | 接受存储在 SharePoint 网站上的数据版本。 |
| **xlListConflictError** | 3 | 如果发生冲突，则引发一个错误。 |
| **xlListConflictRetryAllConflicts** | 1 | 覆盖存储在 SharePoint 网站上的数据版本。 |
