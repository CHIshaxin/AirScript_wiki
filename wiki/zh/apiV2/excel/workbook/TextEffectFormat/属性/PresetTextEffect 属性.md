---
title: "PresetTextEffect (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextEffectFormat/%E5%B1%9E%E6%80%A7/PresetTextEffect%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DtKyrlAn
---

# PresetTextEffect (属性)

返回或设置指定艺术字的样式。**MsoPresetTextEffect** 类型，可读写。

## 说明

该属性的值与**“艺术字库”**对话框内的格式（按从左至右和从上至下的顺序编号）相对应。

| **MsoPresetTextEffect** 可以是下列 **MsoPresetTextEffect** 常量之一。 |
| --- |
| **msoTextEffect1** |
| **msoTextEffect10** |
| **msoTextEffect11** |
| **msoTextEffect12** |
| **msoTextEffect13** |
| **msoTextEffect14** |
| **msoTextEffect15** |
| **msoTextEffect16** |
| **msoTextEffect17** |
| **msoTextEffect18** |
| **msoTextEffect19** |
| **msoTextEffect2** |
| **msoTextEffect20** |
| **msoTextEffect21** |
| **msoTextEffect22** |
| **msoTextEffect23** |
| **msoTextEffect24** |
| **msoTextEffect25** |
| **msoTextEffect26** |
| **msoTextEffect27** |
| **msoTextEffect28** |
| **msoTextEffect29** |
| **msoTextEffect3** |
| **msoTextEffect30** |
| **msoTextEffect4** |
| **msoTextEffect5** |
| **msoTextEffect6** |
| **msoTextEffect7** |
| **msoTextEffect8** |
| **msoTextEffect9** |
| **msoTextEffectMixed** |

设置 **PresetTextEffect**属性会自动设置指定形状的许多其他格式属性。

## 示例

```javascript
/*本示例将第一张工作表上所有艺术字的样式设置为艺术字库对话框中列出的第一个样式。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    for (let s = 1; s <= worksheet.Shapes.Count; s++) {
        if (worksheet.Shapes.Item(s).Type == msoTextEffect) {
            worksheet.Shapes.Item(s).TextEffect.PresetTextEffect = msoTextEffect1
        }
    }
}
```

```javascript
/*此示例显示活动工作表的第三个艺术字的形状是否为艺术字库对话框中列出的第二个样式。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.TextEffect.PresetTextEffect == msoTextEffect2)
}
```
