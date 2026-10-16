---
title: "IgnoreMixedDigits 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SpellingOptions/%E5%B1%9E%E6%80%A7/IgnoreMixedDigits%20%E5%B1%9E%E6%80%A7.html
---
# IgnoreMixedDigits (属性)

如果该属性值为 **False**，则指示 ET 检查混合数字。如果该值为 **True**，则指示 ET 在检查拼写时忽略混合数字。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例中，ET 决定在检查混合数字的拼写时使用何种设置，并通知用户。*/
function test() {
    if (Application.SpellingOptions.IgnoreMixedDigits) {
        console.log("Spelling options for checking mixed digits is disabled.")
    } else {
        console.log("Spelling options for checking mixed digits is enabled.")
    }
}
```

```javascript
/*本示例设置 ET 在检查拼写时检查混合数字。*/
function test() {
    Application.SpellingOptions.IgnoreMixedDigits = false
}
```
