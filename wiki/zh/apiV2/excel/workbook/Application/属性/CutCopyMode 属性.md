---
title: "CutCopyMode (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/CutCopyMode%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Blfm8YOc
---

# CutCopyMode (属性)

返回或设置剪切或复制模式的状态。可为 **True**、**False** 或如下表所示的一个 **XLCutCopyMode** 常量。**Long** 类型，可读写。

## 说明

| 返回值 | 说明 |
|---|---|
| **False** | 不处于剪切或复制模式。 |
| **xlCopy** | 处于复制模式。 |
| **xlCut** | 处于剪切模式。 |

| 设置值 | 说明 |
|---|---|
| **False** | 取消剪切或复制模式并清除移动边框。 |
| **True** | 启动剪切或复制模式并显示移动边框。 |

## 示例

```javascript
/*本示例在消息框中显示剪切或复制模式的状态。*/
function test() {
    switch (Application.CutCopyMode) {
        case 0:
            console.log("Not in Cut or Copy mode")
            break
        case xlCopy:
            console.log("In Copy mode")
            break
        case xlCut:
            console.log("In Cut mode")
            break
    }
}
```

```javascript
/*本示例将A1单元格复制到剪贴板，并判断如果CutCopyMode属性的值为xlCopy，则将其设置为false，并通知用户。*/
function test() {
    Range("A1").Copy()
    if (Application.CutCopyMode == xlCopy) {
        Application.CutCopyMode = false
        console.log("已取消剪切或复制模式并清除移动边框。")
    }
}
```
