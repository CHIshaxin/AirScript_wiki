---
title: "Errors (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Errors/Errors%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CtgUO6yQ
---

# Errors (对象)

---

## 说明

使用 **[Range](./../Range/Range 对象.html)** 对象的 **Errors** 属性可返回 **[Errors](./../Errors/Errors 对象.html)** 对象。

返回 **[Errors](./../Errors/Errors 对象.html)** 对象后，可使用 **[Error](./../Error/Error 对象.html)** 对象的 **Value** 属性检查特定的错误检查条件。

```javascript
/*下例将一个数字作为文本放在单元格 A1 中，然后当单元格 A1 的值包含文本格式的数字时通知用户。*/
function test(){
 //Place a number written as text in cell A1.
    Range("A1").Formula = "'1"

    if(Range("A1").Errors.Item(xlNumberAsText).Value == true) {
        console.log("Cell A1 has a number as text.")
    }
    else {
        console.log("Cell A1 is a number.")
    }

}
```
