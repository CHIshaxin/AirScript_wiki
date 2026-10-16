---
title: "Visible (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comment/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: q_SYlXny
---

# Visible (属性)

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*本示例隐藏单元格 A2 中的批注。*/
function test() {
    Application.Range("A2").Comment.Visible = false
}
```

```javascript
/*本示例显示 A4 单元格的批注是否可见。*/
function test() {
    console.log(Application.Range("A4").Comment.Visible)
}
```
