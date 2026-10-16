---
title: "Version (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/Version%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BM-JUqZk
---

# Version (属性)

返回一个 **String** 值，它代表 ET 版本号。

## 示例

```javascript
/*此示例显示包含 ET 版本号和操作系统名称的消息框。*/
function test() {
    console.log(`Welcome to ET version ${Application.Version} running on ${Application.OperatingSystem}!`)
}
```

```javascript
/*此示例将 ET 的版本转成Number类型，并显示。*/
function test() {
    console.log(Number(Application.Version))
}
```
