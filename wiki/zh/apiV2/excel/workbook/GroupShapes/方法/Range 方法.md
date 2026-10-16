---
title: "Range (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/GroupShapes/%E6%96%B9%E6%B3%95/Range%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DTZZ7LGS
---

# Range (方法)

返回一个 **[ShapeRange](./../../ShapeRange/ShapeRange 对象.html)** 对象，它代表 **[Shapes](./../../Shapes/Shapes 对象.html)** 集合中形状的子集。

## 说明

虽然使用 **Range**属性可返回任意数量的形状，但如果要返回集合的单个成员，用 **Item** 方法更加简单。例如，`Shapes(1)` 比 `Shapes.Range(1)` 简单。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 包含在该区域中的各单个形状。可以是指定形状索引号的整数、指定形状名称的字符串，也可以是包含整数或字符串的数组。 |

## 示例

```javascript
/*此示例设置 myDocument 中第一个形状和第三个形状的填充图案。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick)
}
```

```javascript
/*若要为 Index 指定一个整数或字符串数组，可以使用 Array 函数。例如，以下指令返回用名称指定的两个形状。*/
function test(){
let arShapes = ["Oval 4", "Rectangle 5"]
let objRange = Application.ActiveSheet.Shapes.Range(arShapes)
}
```

```javascript
/*在 ET 中，不能用此属性返回包含工作表上的所有 Shape 对象的 ShapeRange 对象。如果要达到该目的，可用下列代码： */
function test(){
Application.Worksheets.Item(1).Shapes.SelectAll	    // select all shapes
let sr = Selection.ShapeRange	  	    // create ShapeRange
}
```

```javascript
/*此示例设置 myDocument 中形状“Oval 4”和“Rectangle 5”的填充图案。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let arShapes = ["Oval 4", "Rectangle 5"]
let objRange = myDocument.Shapes.Range(arShapes)
objRange.Fill.Patterned(msoPatternHorizontalBrick)
}
```

```javascript
/*此示例设置 myDocument 中第一个形状的填充图案。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let myRange = myDocument.Shapes.Range(1)
myRange.Fill.Patterned(msoPatternHorizontalBrick)
}
```

```javascript
/*此示例创建一个数组，其中包含 myDocument 中所有的自选图形，并用该数组定义一个形状区域，然后在该区域中水平分布所有的形状。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shapes2 = myDocument.Shapes
    numShapes = shapes2.Count
    if(numShapes > 1) {
        let numAutoShapes = 0
        let autoShpArray = []
        for(let i = 1; i <= numShapes; i++) {
            if(shapes2.Item(i).Type == msoAutoShape) {
                autoShpArray[numAutoShapes] = shapes2.Item(i).Name
                numAutoShapes++
            }
        }
        if(numAutoShapes > 1) {
            let asRange = shapes2.Range(autoShpArray)
            asRange.Distribute(msoDistributeHorizontally, false)
        }
    }
}
```
