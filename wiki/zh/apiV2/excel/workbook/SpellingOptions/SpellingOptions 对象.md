---
title: "SpellingOptions (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SpellingOptions/SpellingOptions%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BNA3VDsb
---

# SpellingOptions (对象)

代表工作表的各种拼写检查选项。

## 说明

使用 **[Application](./../Application/Application 对象.html)** 对象的 **[SpellingOptions](./../SpellingOptions/SpellingOptions 对象.html)** 属性可返回一个 **[SpellingOptions](./../SpellingOptions/SpellingOptions 对象.html)** 对象。

一旦返回了 **[SpellingOptions](./../SpellingOptions/SpellingOptions 对象.html)** 对象，您就可以使用下列属性来设置或返回各种拼写检查选项。

- **ArabicModes**
- **DictLang**
- **GermanPostReform**
- **HebrewModes**
- **IgnoreCaps**
- **IgnoreFileNames**
- **IgnoreMixedDigits**
- **KoreanCombineAux**
- **KoreanProcessCompound**
- **KoreanUseAutoChangeList**
- **SuggestMainOnly**
- **UserDict**

## 示例

```javascript
/*下例使用 IgnoreCaps 属性来禁用对全部是大写字母的单词的拼写检查。在本示例中，拼写检查程序发现的是“Testt”，而不是“TESTT”。*/
function test() {
    // 将同一个单词的拼写错误版本全部大写以及混合大小写。
    Range("A1").Formula = "Testt"
    Range("A2").Formula = "TESTT"

    Application.SpellingOptions.SuggestMainOnly = true
    Application.SpellingOptions.IgnoreCaps = true

    // Run a spell check.
    Cells.CheckSpelling()
}
```

```javascript
/*本示例显示 ET 在使用拼写检查时是否忽略大写单词。*/
function test() {
    console.log(Application.SpellingOptions.IgnoreCaps)
}
```
