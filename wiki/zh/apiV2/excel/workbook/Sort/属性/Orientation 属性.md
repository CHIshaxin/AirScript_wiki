---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sort/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BgBiB9wq
---

# Orientation (属性)

指定排序方向。可读/写 **[XlSortOrientation](./../../枚举/XlSortOrientation 枚举.html)** 类型。

## 说明

| XlSortOrientation 可以是下列 **[XlSortOrientation](./../../枚举/XlSortOrientation 枚举.html)** 常量之一: |
| --- |
| **xlSortColumns** |
| **xlSortRows** |

## 示例

```javascript
/*本示例设置第一张工作表执行按列排序。*/
function test() {
    Application.Sheets.Item(1).Sort.Orientation = xlSortColumns
}
```

```javascript
/*本示例判断活动工作表执行的排序方向，并通知用户。*/
function test() {
    if (Application.ActiveSheet.Sort.Orientation == xlSortColumns) {
        console.log("排序时执行按列排序")
    } else {
        console.log("排序时执行按行排序")
    }
}
```
