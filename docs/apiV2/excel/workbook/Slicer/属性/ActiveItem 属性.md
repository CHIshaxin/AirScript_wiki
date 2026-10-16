---
title: "ActiveItem 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/ActiveItem%20%E5%B1%9E%E6%80%A7.html
---
# ActiveItem (属性)

返回一个 **[SlicerItem](./../../SlicerItem/SlicerItem 对象.html)** 对象，该对象代表指定切片器的当前获得焦点的切片器按钮。只读。

## 说明

当指定的切片器有焦点，并且活动的控件是切片器中的按钮时，**ActiveItem** 属性将返回 **[SlicerItem](./../../SlicerItem/SlicerItem 对象.html)** 对象（例如，在此状态下，用户使用键盘切换切片器中的按钮）。

下面的情况下，**ActiveItem** 属性将返回 **Null**：

- 指定的切片器没有焦点（未选中或处于非活动状态）。
- 指定的切片器有焦点，并且选中的是整个切片器本身（例如，用户可在此状态下使用键盘四处移动整个切片器）。
- 指定的切片器有焦点，并且活动的控件是切片器标题中的 **“清除筛选”** 按钮。

## 返回值

SlicerItem

## 示例

```javascript
/*本示例显示活动工作簿上第一个切片器缓存中第一个切片器的焦点中的切片器按钮的名称。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(1).Slicers(1).ActiveItem.Name)
}
```

```javascript
/*本示例取消选择第一张工作簿上第一个切片器缓存的切片器“name”的焦点中的切片器项。*/
function test() {
    let slicer = Workbooks.Item(1).SlicerCaches(1).Slicers("name")
    slicer.ActiveItem.Selected = false
}
```
