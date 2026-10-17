---
title: "FileFormat (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/FileFormat%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D1PmGHgd
---

# FileFormat (属性)

返回工作簿的文件格式和/或类型。XlFileFormat类型，只读。

## 说明

这些常量中的某些可能不可用，这取决于选择或安装的语言支持（例如，美国英语）。

## 示例

```javascript
/*本示例检查当前工作簿文件格式是否为 Excel9795 格式，如果是，则按常规文件格式保存该工作簿。*/
function test() {
    if (Application.ActiveWorkbook.FileFormat == xlExcel9795) {
        Application.ActiveWorkbook.SaveAs(null, xlExcel12)
    }
}
```

```javascript
/*本示例演示当前工作簿FileFormat属性的值。*/
function test() {
    console.log("当前工作簿FileFormat属性的值为：" + Application.ActiveWorkbook.FileFormat)
}
```
