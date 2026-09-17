---
title: "Ignore 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Error/%E5%B1%9E%E6%80%A7/Ignore%20%E5%B1%9E%E6%80%A7.html
---
# Ignore (属性)

允许用户设置或返回某一区域的错误检查选项的状态。如果该值为 **False**，则某一区域的错误检查选项可用。如果该值为 **True**，则禁用某一区域的错误检查选项。**Boolean**类型，可读写。

## 说明

引用 **ErrorCheckingOptions**对象可查看与错误检查选项相关联的索引值列表。

## 示例

```javascript
/*本示例为检查空单元格引用而禁用单元格 A1 中的忽略标记。*/
function test(){
 Range("A1").Select()

    //Determine if empty cell references error checking is on, if not turn it on.
    if(Application.Range("A1").Errors.Item(xlEmptyCellReferences).Ignore == true) {
        Application.Range("A1").Errors.Item(xlEmptyCellReferences).Ignore = false
        console.log("Empty cell references error checking has been enabled for cell A1.")
    }
    else {
        console.log("Empty cell references error checking is already enabled for cell A1.")
    }
}
```
