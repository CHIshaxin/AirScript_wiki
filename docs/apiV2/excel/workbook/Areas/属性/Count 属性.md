---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Areas/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/* 此示例显示 Sheet1 上选定区域中的列数。此示例还将检测选定区域中是否包含多重选定区域，如果包含，则对多重选定区域中每一子区域进行循环*/
function test(){
    Application.Worksheets.Item("Sheet1").Activate()
    let iAreaCount = Application.Selection.Areas.Count

    if(iAreaCount <= 1){
        console.log("The selection contains " + Application.Selection.Columns.Count + " columns.")
    }
    else{
        for(let i = 1; i <= iAreaCount; i++){
            console.log("Area " + i + " of the selection contains " + Application.Selection.Areas.Item(i).Columns.Count + " columns.")
        }
    }
}
```
