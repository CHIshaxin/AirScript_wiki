---
title: "LanguageSettings 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/LanguageSettings%20%E5%B1%9E%E6%80%A7.html
---
# LanguageSettings (属性)

返回一个 **LanguageSettings** 对象，该对象包含 ET 中有关语言设置的信息。只读。

## 示例

```javascript
/*本示例显示安装 ET 时所选语言的语言标识符是否为msoLanguageIDSimplifiedChinese。*/
function test() {
    let languageSet = Application.LanguageSettings
    console.log(languageSet.LanguageID(msoLanguageIDInstall) == msoLanguageIDSimplifiedChinese)
}
```

```javascript
/*本示例显示msoLanguageIDTraditionalChinese是否为编辑的首选语言。*/
function test() {
    console.log(Application.LanguageSettings.LanguagePreferredForEditing(msoLanguageIDTraditionalChinese))
}
```
