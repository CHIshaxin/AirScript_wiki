---
title: "NameLocal (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/NameLocal%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BPJ-F-MY
---

# NameLocal (属性)

以用户语言返回或设置对象的名称。**String** 型，可读写。

## 说明

如果样式为内置样式，则此属性以当前系统所使用的地区语言返回样式的名称。

## 示例

```javascript
/*本示例以用户语言显示活动工作簿中索引号为4的自定义名称。*/
function test() {
    console.log(`索引号为4的名称为 ${ActiveWorkbook.Names.Item(4).NameLocal}`)
}
```

```javascript
/*本示例以用户语言显示活动工作簿中索引号为3的名称是否为“test”。*/
function test() {
    console.log(ActiveWorkbook.Names.Item(3).NameLocal == "test")
}
```
