---
title: "Save 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E6%96%B9%E6%B3%95/Save%20%E6%96%B9%E6%B3%95.html
---
# Save (方法)

保存对指定工作簿所做的更改。

## 说明

要打开工作簿文件，请使用**Open**方法。

要将工作簿标记为已保存，但不将其写入磁盘，请将它的**Saved** 属性设置为 **True**。

首次保存工作簿时，请使用 SaveAs 方法指定文件名。

## 示例

```javascript
/*本示例保存当前活动工作簿。*/
function test() {
    Application.ActiveWorkbook.Save()
}
```

```javascript
/*本示例保存所有打开的工作簿，然后关闭 ET。*/
function test() {
    for (let w = 1; w <= Application.Workbooks.Count; w++) {
        Application.Workbooks.Item(w).Save()
    }
    Application.Quit()
}
```
