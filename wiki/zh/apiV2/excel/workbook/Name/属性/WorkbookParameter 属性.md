---
title: "WorkbookParameter (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E5%B1%9E%E6%80%A7/WorkbookParameter%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BWLOrcPH
---

# WorkbookParameter (属性)

以工作簿参数的形式返回或设置指定的 **[Name](./../../Name/Name 对象.html)** 对象。可读/写。

## 返回值

Boolean

## 示例

```javascript
/*本示例显示索引为2的名称是否被定义为工作簿参数。*/
function test() {
    console.log(ActiveWorkbook.Names.Item(2).WorkbookParameter)
}
```

```javascript
/*本示例将活动工作簿索引为1的名称设置为工作簿参数。*/
function test() {
    ActiveWorkbook.Names.Item(1).WorkbookParameter = true
}
```
