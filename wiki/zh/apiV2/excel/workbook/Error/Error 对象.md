---
title: "Error (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Error/Error%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: twpzOqt7
---

# Error (对象)

代表区域的电子表格错误。

## 说明

使用 **[Errors](./../Errors/Errors 对象.html)** 对象的 **Item** 属性可返回 **[Error](./../Error/Error 对象.html)** 对象。

返回 **[Error](./../Error/Error 对象.html)** 对象后，可结合使用 **Value** 和 **Errors**属性检查某个特定错误检查选项是否已启用。

| 注释 |
| --- |
| 不要将 **[Error](./../Error/Error 对象.html)** 对象与 Visual Basic 的错误处理功能混淆。 |

```javascript
/*下例中，在引用空单元格的单元格 A1 中创建一个公式，然后使用 Item(index)（其中 index 用于标识错误类型）显示描述错误情况的消息。*/
function test(){
 let rngFormula = Application.Range("A1")

    //Place a formula referencing empty cells.
    Range("A1").Formula = "=A2+A3"
    Application.ErrorCheckingOptions.EmptyCellReferences = true

    //Perform check to see if EmptyCellReferences check is on.
    if(rngFormula.Errors.Item(xlEmptyCellReferences).Value == true) {
        console.log("The empty cell references error checking feature is enabled.")
    }
    else {
        console.log("The empty cell references error checking feature is not on.")
    }
}
```
