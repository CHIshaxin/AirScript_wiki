---
title: "邮件 API"
source: https://airsheet.wps.cn/docs/api/advanced/SMTP.html
---
# 邮件 API

通过外部邮件服务发送邮件。

### 快速使用

```javascript
// 登录
let mailer = SMTP.login({
    host: "smtp.example.com", // 域名
    port: 465, // 端口
    secure: true, // TLS
    username: "sender@example.com", // 账户名
    password: "Pa55W0rd" // 密码
})
// 客户端发送邮件
mailer.send({
    from: "sender@example.com", // 发件人
    to: "reciever@example.com", // 收件人
    subject: "this is subject.", // 主题
    text: "this is text.", // 文本
    html: `<p> this is html </p>` // HTML代码
})
// 支持指定昵称
mailer.send({
    from: "管理员 <admin@example.com>",
    to: "接受者 <username@example.com>",
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
// 支持发送多个邮箱
mailer.send({
    from: "管理员 <admin@example.com>",
    to: ["username1@example.com","接受者2 <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

## SMTP

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| login(argvs) | Mailer | 登录并返回邮件发送者 |

## login(argvs)

登录并返回邮件发送对象

```javascript
//  登录qq邮箱
let mailer = SMTP.login({
    host: "smtp.qq.com", // QQ 的SMTP服务器的域名
    port: 465,
    username: "1000000000@qq.com", // qq 邮箱地址
    password: "xxxxxxxxxxxx", // qq邮箱的SMTP密码，非qq密码
    secure: true
});
```

### 参数

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| argvs | LoginArgvs | undefined | true | 一个JavaScript对象，用于配置SMTP的参数，如下所示 |

### LoginArgvs

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| host | string | undefined | true | 邮箱服务器域名 |
| port | number | undefined | true | SMTP服务端口，当host为undefined时，取默认值，默认值由secure决定，当secure是false时默认值为587，当secure是true时默认值为465。 |
| secure | boolean | undefined | true | 是否使用TLS连接服务器，在大多数情况下，如果要连接到端口465，请将此值设置为true；如果要连接到端口587或25，请将此值设置为false。 |
| username | string | undefined | true | 用于身份验证的账户名 |
| password | string | undefined | true | 用于身份验证的密码 |
| timeout | number | 10000 | false | 等待建立连接的时间，单位毫秒(ms) |

### 返回值

Mailer- 邮件发送者

## Mailer

由 login(argvs)创建的对象，用于发送邮件

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| send(message) | undefined | 发送邮件 |

## send(message)

发送邮件

```javascript
mailer.send({
    from: ["Administrator <admin@example.com>"],
    to: ["username@example.com", "UserName <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

### 参数

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| message | messageArgvs | undefined | true | 一个JavaScript对象，要发送的邮件内容，如下所示 |

### messageArgvs

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| from | string | undefined | true | 发件人的电子邮箱地址 |
| to | string / string[] | undefined | true | 收件人的电子邮箱地址 |
| subject | string | undefined | true | 电子邮件的主题 |
| text | string | undefined | true | 电子邮件显示的文本 |
| html | string | undefined | false | 电子邮件的HTML代码 |
