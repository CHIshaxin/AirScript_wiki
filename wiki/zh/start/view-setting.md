---
title: "配置视图"
source: https://airsheet.wps.cn/docs/start/view-setting.html
updated: 1789700916000
hash: DyeUVuk_
---

# 配置视图

为方便用户将写的脚本交付给无开发经验的第三方使用，我们推出了参数可视化配置。用户通过点击脚本工具栏`视图配置`，可以打开参数编排工具。通过配置`form`表单，结合脚本便可生成快捷操作工具。

## 视图与脚本的关联

- 视图配置的表单参数，会在脚本编辑器中将参数挂载在`Context.argv`下，编程的过程中，可直接引用此变量。`Context`作为全局变量，类似`Application`。
- 视图配置的参数会实时在代码编辑器内生成语法提示，输入Context.argv，就可看到已配置的参数。
- 预览视图，点击视图内`运行脚本`按钮，将携带表单输入参数运行脚本，同时打印日志复用代码编辑器的日志框。

## 简单使用案例

#### 1、配置参数

打开视图配置，添加以下参数：

| 参数名称 | 变量名称 | 变量类型 | 输入提示 |
|---|---|---|---|
| 选择表 | sheet | string | 输入表名：“sheet1” |
| 选择单元格 | range | string | 输入单元格：“A1” |
| 输入设置值 | value | string | 输入要设置的值 |

#### 2.输入代码

```js
Application.Sheets.SelectByName(Context.argv.sheet).Activate()
Application.Range(Context.argv.range).Value=Context.argv.value
```

#### 3.编辑效果
