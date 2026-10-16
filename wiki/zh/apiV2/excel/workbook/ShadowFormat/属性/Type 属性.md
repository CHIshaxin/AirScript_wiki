---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShadowFormat/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BNr0GKFi
---

# Type (属性)

返回或设置一个代表阴影格式类型的 **MsoShadowType** 值。

## 示例

```javascript
/*本示例显示第一张工作表中第二个形状阴影的格式类型是否为msoShadow17。*/
function test() {
    let shadow = Application.Worksheets.Item(1).Shapes.Item(2).Shadow
    console.log(shadow.Type == msoShadow17)
}
```

```javascript
/*本示例将活动工作表中第一个形状阴影的格式类型改为msoShadow16。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(1).Shadow
    shadow.Type = msoShadow16
}
```
