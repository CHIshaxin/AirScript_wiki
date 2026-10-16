---
title: "工作表(Sheet)"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Sheet.html
---
# Sheet

工作簿（Workbook）中单个工作表(Sheet)对象

Sheet 对象的具体属性和方法请参阅下方的列表。

#### 属性列表

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Id | String | 该工作表的 Id |
| Name | String | 该工作表的名称 |
| Index | Number | 该工作表在所有工作表的索引值 |
| Cells | [Range](Range.md) | 该工作表上所有单元格的集合 |
| Columns | [Range](Range.md) | 该工作表上所有列的集合 |
| Rows | [Range](Range.md) | 该工作表上所有行的集合 |
| UsedRange | [Range](Range.md) | 该工作表的使用范围 |
| Visible | Boolean | 该工作表是否可见 |
| Type | String | 该工作表的类型 |
| Hyperlinks | Hyperlinks | 该工作表上所有超链接的集合 |
| Shapes | Shapes | 该工作表上所有图形的集合 |
| Sort | Sort | 该工作表上排序对象 |

#### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Range() | [Range](Range.md) | 一个单元格或单元格区域 |
| Cells() | [Range](Range.md) | 该工作表上的某个单元格 |
| Activate() | undefined | 切换(激活)工作表 |
| Move() | undefined | 移动工作表 |
| Delete() | undefined | 删除工作表 |

## Id

获取工作表 Id

#### 数据类型

String - 工作表 Id

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表的id
console.log(sheet.Id)
```

## Name

设置/获取 工作表名称

#### 数据类型

String - 该工作表在所有工作表的名称

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表的名称
console.log(sheet.Name) // Sheet2

// 将当前工作表的名称改为 WPS WebOffice
sheet.Name = 'WPS WebOffice'
```

## Index

工作表的 index,即该工作表在所有工作表的索引值

#### 数据类型

String - 该工作表在所有工作表的索引值

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表的index
console.log(sheet.Index) // 1
```

## Cells

工作表上所有单元格的集合，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### 数据类型

[Range](Range.md) - 工作表上所有单元格的集合

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印活动工作表的全部单元格地址
console.log(sheet.Cells.Address()) // $A$1:$XFD$1048576
```

## Columns

工作表上的所有列，返回的是一个 Range 对象，可参考使用 [Range](Range.md)

#### 数据类型

[Range](Range.md) - 工作表上的所有列

#### 示例

```js
// 打印该工作表列的数量
const sheet = Application.ActiveSheet
console.log(sheet.Columns.Count) //16384
```

## Rows

工作表上的行，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### 数据类型

[Range](Range.md) - 工作表上的所有列

#### 示例

```js
const sheet = Application.ActiveSheet
console.log(sheet.Rows.Count) // 1048576
```

## UsedRange

工作表激活的区域,即是工作表实际使用到的区域，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### 数据类型

[Range](Range.md) - 工作表里的一个单元格或单元格区域

#### 示例

```js
const UsedRange = Application.ActiveSheet.UsedRange
// 打印激活区域的范围,此处假设激活区域为10*10的单元格
console.log(
  UsedRange.Row,
  UsedRange.RowEnd,
  UsedRange.Column,
  UsedRange.ColumnEnd
) // 1,10,1,10
```

## Visible

显示/隐藏 工作表

#### 数据类型

Boolean - 工作表是否可见

#### 示例

```js
const sheet = Application.ActiveSheet
// 隐藏工作表
sheet.Visible = false
// 取消工作表隐藏
sheet.Visible = true
```

## Type

工作表类型

#### 数据类型

[Enum.XlSheetType](Enum.md#xlsheettype) - 工作表的类型

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表的类型
console.log(sheet.Type) //xlWorksheet
```

## AutoFilter

当前工作表的自动筛选对象，该对象内封装了当前工作表的所有筛选对象集合，以及应用和取消筛选条件的方法

#### 数据类型

[AutoFilter](AutoFilter.md) - 自动筛选对象

#### 示例

```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取筛选对象集合
const filters = autoFilter.Filters
// 获取当前工作表所有筛选对象的数量
const count = filters.Count
console.log(count)
```

## Hyperlinks

工作表上的所有超链接的集合

#### 数据类型

Hyperlinks集合 - 所有超链接的集合对象

#### 示例

```js
// 打印超链接集合里超链接对象的个数
const hyperlinks = Application.ActiveSheet.Hyperlinks
console.log(hyperlinks.Count) // 3
```

### Hyperlinks.Count

集合中超链接的数量

#### 数据类型

Number - 集合中超链接的数量

#### 示例

```js
// 打印超链接集合里超链接对象的个数
const hyperlinks = Application.ActiveSheet.Hyperlinks
console.log(hyperlinks.Count) // 3
```

### Hyperlinks.Item()

获取超链接对象

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Index | Number |  | 是 | 从 1 开始 |

#### 返回类型

[Hyperlink](Hyperlink.md) - 超链接对象

#### 示例

```js
const hyperlinks = Application.ActiveSheet.Hyperlinks
if (hyperlinks.Count <= 0) {
  throw new Error('当前文档没有超链接')
}
// 打印超链接的地址和文本
console.log(hyperlinks.Item(1).TextToDisplay) // 超链接的显示内容
console.log(hyperlinks.Item(1).Address) // 点击跳转的超链接
```

## Shapes

当前工作表上的所有 Shape 对象的集合

#### 数据类型

Shapes集合对象 - 该工作表上的所有 Shape 对象的集合

#### 示例

```js
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
const chartEnum = Application.Enum.XlChartType.xlColumnClustered //簇状柱形图
// 在图形对象集合中添加 300 * 300 的簇状柱形图
shapes.AddChart2(340, chartEnum, 0, 0, 300, 300)
```

### Shapes.GetActiveShapeImg()

获取激活单元格的图片数据

#### 返回类型

String - 图片原图下载链接

#### 示例

```js
// 假如A1有图片
Application.Range('A1').Select()
const imgUrl = Application.ActiveSheet.Shapes.GetActiveShapeImg()
console.log(imgUrl) // https://imageUrl 如果没有图片则返回undefined
```

### Shapes.Item(Index)

代表绘图层中的对象，例如自选图形、任意多边形、OLE 对象或图片

#### 参数

| 参数名 | 数据类型 | 默认值 | 可选 | 简介 |
| --- | --- | --- | --- | --- |
| index | Number |  | 否 | 对象的索引 |

#### 返回类型

[Shape](Shape.md) - 绘图层的一个对象,自选图形、任意多边形、OLE 对象或图片

#### 示例

```js
const shapes = Application.ActiveSheet.Shapes
// 打印第一个图形对象的ID
console.log(shapes.Item(1).ID)
```

### Shapes.Count

图形的数量

#### 数据类型

Nunber - 工作表的图形对象的数量

#### 示例

```js
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
// 打印当前图形的数量
console.log(shapes.Count)
```

### Shapes.AddChart2()

添加图形

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Style | String |  | 否 | 指定图表样式 |
| XlChartType | Enum |  | 否 | 指定图表类型，对应 [Enum.XlChartType](Enum.md#xlcharttype) |
| Left | Number |  | 否 | 指定新建图表的左边距，单位 px |
| Top | Number |  | 否 | 指定新建图表的上边距，单位 px |
| Width | Number |  | 否 | 指定新建图表的宽度，单位 px |
| Height | Number |  | 否 | 指定新建图表的高度，单位 px |

#### 返回类型

[Shape](Shape.md) - 添加的图形对象

#### 示例

```js
const chartEnum = Application.Enum.XlChartType.xlColumnClustered //簇状柱形图的枚举
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
// 在图形对象集合中添加 300 * 300 的簇状柱形图
shapes.AddChart2(340, chartEnum, 0, 0, 300, 300)
```

### Shapes.AddPicture()

向表格中插入浮动图片

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| FileName | String |  | 是 | 要插入的图片（可以是图片数据的 Base64 字符串或者 URL 图片地址） |
| LinkToFile | Enum |  | 是 | 暂不支持，请填`0` |
| SaveWithDocument | Enum |  | 是 | 暂不支持，请填`0` |
| Left | Number |  | 是 | 图片左边缘相对于表格左边缘的位置 |
| Top | Number |  | 是 | 图片上边缘相对于表格片上边缘的位置 |
| Width | Number |  | 否 | 图片的宽度 |
| Height | Number |  | 否 | 图片的高度 |

#### 示例

```js
// 获取图形对象
const shapes = ActiveSheet.Shapes
// 插入浮动图片
shapes.AddPicture({
  FileName:
    'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAACH1BMVEUAAAA9kP8mpv9Fv/8QT94TXeQAfv8trP0xr/0VUdoAcv8LgPs1sf4AgP4QUuURTuEVl/ksq/wAg/84tf4NUt80sf0bnPsPUuAAYfsAgP4UTtYAY/8GXPQRkv1Jxv8AZP86tv9Bvv9Gv/9Atv8KiPUOVesAZ+wbS8wUl/kjo/sAYv8PVekjovxBu/8AY/8trPxCu/8AW/AAcvAxsP0npvsYTdElo/oQT9o+uP8AYv8KVOIHdugAYfwAfPwbfe49uf4XmfkcdOopo/oAYv86tf4AYf8TT9cZaeYVTdUJVOUcnvsAYv8Agv9Iwv8AYv9Jwf8bSs0gofoAY/8Ag/9Gv/8en/oAYf9Fv/8AWuQZS80AdO5BvP8AYv9DvP8Umfo9uv8LVOUAgf8amfk/uf8AY/8Ag/8Ahf8TUdgVT9UNWOknov8trPwxr/wmpvspqfs1sv0io/s5tf0AgP8bnfoMV+4AYPoAYv8AYf4AYfwAe/sen/oAd/gAXvYAff0LW/MLWfEOUN0PT9oSTtZAu/49uP0qqfweoPoXmvkAdPUAXPIAcfEKUuIWTNIXS88KXPYAWusOVeoPVOkPUuYQUeQQT+IMUeAUTdMZSs09uP4Aa+kAWOgAaeQIVOQIU+QAV+MRTd8STd0RTtgSl/kAW+8NWO8Abu0AbuwBZNsAg/8AYPwAc/MIVecAZuEAVeAAWt8AZN0JgvAIfOsJV+sTVuIAU9qil9AQAAAAa3RSTlMAAwb++A3+bkkkGxL8+Pf36tnTubCYjol4dmVjRjEvLysZFA7+/Pv6+fj39/b19fPw8O/u7Ovq6Ofn5+Xj4uDa2trZ1dPS0s/Ny7+9vLi4trapoqGgoJ2XkI+KgoJycGRhX1BNS0pHQj06IcCB3jkAAAIBSURBVDjLdZCHctpAFEWfCMU1ce8tTu+99957770nJFGwkE1sgyMTAginYbkAIQSD494+0G+ltTAacWZnNLt7dO/Og3Tqzl2og8wsvroAudEI+iy5s+6LzMZ7Bp1r5sFOQegRenAJwoEnoCXr6Mc0ihvSrhuK32vJK21Mva00760OWyz0KZYNbzJQmKXUF1o7rZ24rGngweETIGOwbH2nw/ZDUaM6grJVHzSs3R+NoqDy+lKLuwWX2618r03/QRalhNzbewo+qZyuZX4TVMFQubm3f3xy5VeZg49xsL8IC+mUH+1zeLyJ3pHkRCwWK7hvIGc/CYrwvIhtbXZ4hgYwJLm0LFf56QdhDRAq8wdZMypD3kT/SBJo6nfCMmWTc2WQlUO8GDIn2GSAUnuWGERJpIS4LR4HlWdFNITua9pl5s3h5vJ81tzscAAhuyTc3kZQ51CxqasvMmDGENyZyo1BX2iqLRAI0L7qvU2fv6Ex5mVbgana4XIGfb7QsI0KNcc4O4/Gv67I6Jjn5cmOvy4nGuFwaFipqFjN2YlBQkZn/N0dxAhiSfjUKzqHyxwq/0lIX0T0d6OAGb7dD0Gl/gzH8UqN6JcNp7HcBPN5esTO8U2oSCIpcZVkgwamepedxxpJROP4C9DBdHc9vlWSxG1VDOiTc53jpRW3TJCZ+vMXNeWzDz4DoNZyqecAAAAASUVORK5CYII=',
  LinkToFile: 0,
  SaveWithDocument: 0,
  Left: 30, // 图片距离左边位置
  Top: 30, // 图片距离顶部位置
  Width: 32, // 图片宽度
  Height: 32 // 图片高度
})
```

## Sort

排序对象，设置好目标区域后即可进行排序操作

#### 数据类型

[Sort](Sort.md) - 排序对象

#### 示例

```js
// 对 C 列降序排序并且对 D 列做升序排序

//获取当前表格区域
const range = ActiveSheet.UsedRange
//获取到排序对象
const sort = ActiveSheet.Sort
//获取排序范围
const sortFields = sort.SortFields
//清除之前的范围
sortFields.Clear()
//基于C列降序排序, xlSortOnValues代表按值排序, xlDescending代表降序排序
sortFields.Add(Range('C:C').Item(1, 1), xlSortOnValues, xlDescending)
//基于D列升序排序, xlSortOnValues代表按值排序, xlAscending代表升序排序
sortFields.Add(Range('D:D').Item(1, 1), xlSortOnValues, xlAscending)
//设置是否包含表头参数，xlGuess为自动，xlYes为包含表头，xlNo为不包含表头。默认设置为xlGuess
sort.Header = xlGuess
//设置是否大小写敏感，true为区分大小写，false为不区分大小写，默认设置false
sort.MatchCase = false
//设置中文排序方法，xlPinYin为拼音排序，xlStroke为比划数排序。默认设置为xlPinYin
sort.SortMethod = xlPinYin
//设置排序的方法，xlSortColumns为按列排序，xlSortRows为按行排序，默认设置为xlSortColumns
sort.Orientation = xlSortColumns
//排序前必须设置SetRange
sort.SetRange(range)
//开始排序
sort.Apply()
```

## Range()

一个单元格或单元格区域，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### 数据类型

[Range](Range.md) - 工作表里的一个单元格或单元格区域

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表D2单元格的内容
console.log(sheet.Range('D2').Text) // D2

// 修改当前活动工作表D2单元格的内容
sheet.Range('D2').Value = 'this is D2'
```

## Cells()

选择工作表上的某个单元格，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### 返回类型

[Range](Range.md) - 工作表上的某个单元格

#### 示例

```js
// 打印活动工作表的第一个单元格地址
console.log(Application.Cells(1).Address()) // $A$1
console.log(Application.Cells(2).Address()) // $A$2
```

## Activate()

激活工作表

#### 示例

```js
const sheet = Application.Sheets.Item(1)
// 激活第一个工作表，此时的Application.ActiveSheet.Range和此时的Application.Range都指向第一个工作簿
sheet.Activate()
// 修改了第一个工作表的A1单元格
Application.Range('A1').Value = 'foo'
```

## Move()

移动工作表

#### 参数

**两个参数互斥**

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Before | number | null | 否 | 验将放置移动的工作表之前的工作表 ID。如果指定 After ，则不能指定 Before。 |
| After | number | null | 否 | 将放置移动的工作表后的工作表 ID。如果指定 Before ，则不能指定 After |

#### 示例

```js
// 将当前工作表移动到第二个工作表之后
const sheet = Application.ActiveSheet
sheet.Move({
  Before: null,
  After: Application.Sheets(2).Id
})
```

## Delete()

删除工作表

#### 返回值

undefined

#### 示例

```js
// 删除名称为“Sheet2”的工作表
Application.Sheets.Item('Sheet2').Delete()
```
