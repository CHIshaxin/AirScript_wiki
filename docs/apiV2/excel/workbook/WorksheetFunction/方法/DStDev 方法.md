---
title: "DStDev 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/DStDev%20%E6%96%B9%E6%B3%95.html
---
# DStDev (方法)

将列表或数据库的列中符合指定条件的数字作为样本，估算总体的标准偏差。

## 说明

- 由于在单元格中键入文本或值时等号用来指示一个公式，因此 ET 将计算键入内容；但是，这可能会导致产生意外的筛选结果。要为文本或值指示等号比较运算符，请在条件区域的相应单元格中将条件作为字符串表达式键入： **="=entry"** 其中，entry 是要查找的文本或值。例如：

| 在单元格中键入的内容 | ET 计算和显示的内容 |
| --- | --- |
| ="=Davolio" | =Davolio |
| ="=3000" | =3000 |

- ET 在筛选文本数据时不区分大小写。但是，可使用公式执行区分大小写的搜索。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | Range | 必填 | Database - 构成列表或数据库的单元格区域。数据库是相关数据的列表，其中包含相关信息的行为记录，而包含数据的列为字段。列表的第一行包含每一列的标签。 |
| Arg2 | any | 必填 | Field - 指示函数所使用的列。输入使用双引号括起来的列标签，如“使用年数”或“产量”，或者输入代表列表中列位置的数字（不带引号）：1 代表第一列，2 代表第二列，依此类推。 |
| Arg3 | any | 必填 | Criteria - 包含指定条件的单元格区域。只要参数 criteria 至少包含一个列标签和列标签下方用于为该列指定条件的一个单元格，就可以为它使用任意区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例为D1等单元格分别赋值，使用DStDev方法计算从D3:F6相应的列中提取符合指定条件的数字作为样本，估算总体的标准偏差，并将结果赋值于E2单元格。*/
function test() {
    Range("D1").Value2 = "分数"
    Range("D2").Value2 = ">=70"
    Range("E1").Value2 = "标准偏差"
    Range("F1").Value2 = "分数"
    Range("F2").Value2 = "<=100"
    Range("D3").Value2 = "专业"
    Range("D4").Value2 = "软件工程"
    Range("D5").Value2 = "软件工程"
    Range("D6").Value2 = "软件工程"
    Range("E3").Value2 = "分数"
    Range("E4").Value2 = 67
    Range("E5").Value2 = 85
    Range("E6").Value2 = 86
    Range("F3").Value2 = "姓名"
    Range("F4").Value2 = "小明"
    Range("F5").Value2 = "李四"
    Range("F6").Value2 = "张三"
    Range("E2").Value2 = Application.WorksheetFunction.DStDev(Range("D3:F6"), 2, Range("D1:F2"))
}
```

```javascript
/*本示例为B1等单元格分别赋值，并使用DStDev方法计算从B1:C4相应的列中提取符合指定条件的数字作为样本，估算总体的标准偏差。*/
function test() {
    Range("B1").Value2 = "水果"
    Range("B2").Value2 = "西瓜"
    Range("B3").Value2 = "苹果"
    Range("B4").Value2 = "葡萄"
    Range("C1").Value2 = "单价"
    Range("C2").Value2 = 2.4
    Range("C3").Value2 = 8
    Range("C4").Value2 = 14
    Range("D1").Value2 = "单价"
    Range("D2").Value2 = ">3.5"
    Range("E1").Value2 = "单价"
    Range("E2").Value2 = "<20"
    let det1 = Application.WorksheetFunction.DStDev(Range("B1:C4"), Range("C1").Value2, Range("D1:E2"))
    console.log(det1)
}
```
