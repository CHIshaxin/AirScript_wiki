---
title: "IgnoreCaps 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SpellingOptions/%E5%B1%9E%E6%80%A7/IgnoreCaps%20%E5%B1%9E%E6%80%A7.html
---
# IgnoreCaps (属性)

如果该属性值为 **False**，则指示 ET 检查大写单词。如果该值为 **True**，则指示 ET 在使用拼写检查时忽略大写单词。可读/写 **Boolean** 类型。

## 示例

```javascript
/*在本示例中，ET 将确定检查大写单词的拼写时使用何种设置，并通知给用户。*/
function test() {
    if (Application.SpellingOptions.IgnoreCaps) {
        console.log("Spelling options for checking uppercase words is disabled.")
    } else {
        console.log("Spelling options for checking uppercase words is enabled.")
    }
}
```

```javascript
/*本示例设置 ET 在使用拼写检查时检查大写单词。*/
function test() {
    Application.SpellingOptions.IgnoreCaps = false
}
```
