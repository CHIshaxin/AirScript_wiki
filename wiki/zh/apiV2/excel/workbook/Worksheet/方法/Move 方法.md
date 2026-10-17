---
title: "Move (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/Move%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Ckeue69L
---

# Move (方法)

将工作表移到工作簿中的其他位置。

## 说明

如果既不指定 *Before* 也不指定 *After*，ET 将新建一个工作簿，其中包含所移动的工作表。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Before | any | 可选 | 在其之前放置移动工作表的工作表。如果指定了 After，则不能指定 Before。 |
| After | any | 可选 | 在其之后放置移动工作表的工作表。如果指定了 Before，则不能指定 After。 |

## 示例

```javascript
/*此示例将当前活动工作簿的 Sheet1 移到 Sheet3 之后。*/
function test() {
    Application.Worksheets.Item("Sheet1").Move(null, Application.Worksheets.Item("Sheet3"))
}
```

```javascript
/*本示例判断如果工作表数量大于1，则将 Sheet2 移到 Sheet1 之前。*/
function test() {
    if (Worksheets.Count > 1) {
        Worksheets.Item("Sheet2").Move(Application.Worksheets.Item("Sheet1"))
    }
}
```
