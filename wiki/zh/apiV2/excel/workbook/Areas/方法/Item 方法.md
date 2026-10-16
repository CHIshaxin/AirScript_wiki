---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Areas/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: rKTkgMSi
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Index | int4 | 必填 | 对象的索引号。 |

## 示例

```javascript
/* 此示例检查当前选定区域是否为多重选定区域，如果是，则清除其中的第一个子区域的内容*/
function test() {
  if(Application.Selection.Areas.Count != 1){
      Application.Selection.Areas.Item(1).Clear()
  }
}
```
