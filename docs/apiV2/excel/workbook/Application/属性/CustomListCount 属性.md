---
title: "CustomListCount 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/CustomListCount%20%E5%B1%9E%E6%80%A7.html
---
# CustomListCount (属性)

返回已定义的自定义序列的数目（包括内置序列）。**Long** 类型，只读。

## 示例

```javascript
/*本示例显示当前已定义的自定义序列的数量。*/
function test() {
    console.log(`There are currently ${Application.CustomListCount} defined custom lists`)
}
```

```javascript
/*本示例判断是否存在已定义的自定义序列，在通知用户后，增加一个新的自定义序列。*/
function test() {
    if (Application.CustomListCount > 0) {
        console.log("已定义的自定义序列的数目为：" + Application.CustomListCount)
    }
    Application.AddCustomList(["Year", "Month", "Day"])
    console.log("已定义的自定义序列的数目为：" + Application.CustomListCount)
}
```
