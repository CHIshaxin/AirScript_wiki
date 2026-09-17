---
title: "EnablePivotTable 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/EnablePivotTable%20%E5%B1%9E%E6%80%A7.html
---
# EnablePivotTable (属性)

如果当仅限于用户界面保护处于打开状态时，启用数据透视表控件和操作，则该属性值为 **True**。**Boolean** 类型，可读写。

## 说明

该属性可应用于每一张工作表，而不随工作表或会话一起保存。

在数据透视表的下方或右侧必须有足够的未锁定单元格，这样 ET 才能重新计算并显示该数据透视表。

## 示例

```javascript
/*本示例在受保护的工作表上启用数据透视表控件。*/
function test() {
    Application.ActiveSheet.EnablePivotTable = true
    Application.ActiveSheet.Protect(null, null, true, null, true)
}
```

```javascript
/*本示例演示工作表 Sheet1 是否启用数据透视表控件。*/
function test() {
    console.log(Application.Sheets.Item("Sheet1").EnablePivotTable)
}
```
