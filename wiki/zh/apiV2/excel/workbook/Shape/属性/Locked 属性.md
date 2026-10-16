---
title: "Locked (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Locked%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Djo5vldL
---

# Locked (属性)

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

## 说明

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**。

## 示例

```javascript
/*本示例显示活动工作表第一个形状的锁定状态。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(1).Locked == true) {
        console.log("形状已锁定")
    } else {
        console.log("形状未锁定")
    }
}
```

```javascript
/*本示例将第一张工作表中的所有形状锁定状态改为未锁定。*/
function test() {
    let shapes = Worksheets.Item(1).Shapes
    for (let i = 1; i <= shapes.Count; i++) {
        shapes.Item(i).Locked = false
    }
}
```
