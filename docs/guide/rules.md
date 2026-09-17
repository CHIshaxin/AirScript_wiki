---
title: "脚本语言"
source: https://airsheet.wps.cn/docs/guide/rules.html
---
# 脚本语言

AirScript脚本采用标准 [JavaScript](https://developer.mozilla.org/zh-CN/docs/Learn/Getting_started_with_the_web/JavaScript_basics)语言进行编写，支持大部分ES6语法。

## 内置全局对象

#### 文档OpenApi对象 `Application`

Application为文档OpenApi对象，基于不同的金山文档类型，会生成不同的Application对象属性，在代码编辑器内可直接引用。

#### 脚本上下文对象 `Context`

Context为脚本运行上下文对象。负责挂载一些与脚本运行相关的属性，如`视图`、`配置`等产生的相关内容，在代码编辑器内可直接引用，以方便整体应用的构造。

## 不支持的语法

### Class

```javascript
class C{
    private a; // Unexpected token
    public b(){
        return this.a;
    }
}
```

### Object里直接定义方法

```javascript
let obj = {
    sayHello(){ //Unexpected token
        console.log("hello")
    }
}
obj.sayHello()
```

### import、export

```javascript
import os from 'os' // SyntaxError
export default os // SyntaxError
```

### ?.

```javascript
function f(argv){
    console.log(argv?.toString()) // Unexpected token
}
```

### await

```javascript
async function f() {
    await new Promise((resolve)=>{ // Unexpected token
        resolve("")
    })
}
```

### yield

```javascript
function* foo(index) {
  while (index < 2) {
    yield index; // Unexpected token
    index++;
  }
}

const iterator = foo(0);

console.log(iterator.next().value);

console.log(iterator.next().value);
```
