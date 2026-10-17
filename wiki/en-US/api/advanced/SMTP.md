---
title: "Email API"
source: https://airsheet.wps.cn/docs/en-US/api/advanced/SMTP.html
updated: 1789700916000
hash: CLNBttb0
---

# Email API

Send email through an external mail service.

### Quick start

```javascript
// Login
let mailer = SMTP.login({
    host: "smtp.example.com", // Hostname
    port: 465, // Port
    secure: true, // TLS
    username: "sender@example.com", // Account
    password: "Pa55W0rd" // Password
})
// Send mail
mailer.send({
    from: "sender@example.com", // From
    to: "reciever@example.com", // To
    subject: "this is subject.", // Subject
    text: "this is text.", // Plain text
    html: `<p> this is html </p>` // HTML
})
// Display name supported
mailer.send({
    from: "Administrator <admin@example.com>",
    to: "Recipient <username@example.com>",
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
// Multiple recipients
mailer.send({
    from: "Administrator <admin@example.com>",
    to: ["username1@example.com","Recipient 2 <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

## SMTP

### Methods

| Method | Return type | Summary |
| --- | --- | --- |
| [login(argvs)](#login-argvs) | [Mailer](#mailer) | Logs in and returns a mail sender |

## login(argvs)

Logs in and returns a mail sender object.

```javascript
// GMail login
let mailer = SMTP.login({
    host: "smtp.gmail.com", // GMail SMTP host
    port: 465,
    username: "1000000000@gmail.com", // Gmail address
    password: "xxxxxxxxxxxx", // SMTP password, not the GMail account password
    secure: true
});
```

### Parameters

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| argvs | LoginArgvs | undefined | true | JavaScript object with SMTP settings, see below |

### LoginArgvs

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| host | string | undefined | true | Mail server hostname |
| port | number | undefined | true | SMTP port. When `host` is undefined, default depends on `secure`: false defaults to 587, true defaults to 465. |
| secure | boolean | undefined | true | Whether to use TLS. Usually set true for port 465; set false for 587 or 25. |
| username | string | undefined | true | Account name for authentication |
| password | string | undefined | true | Password for authentication |
| timeout | number | 10000 | false | Connection timeout in milliseconds (ms) |

### Return value

[Mailer](#mailer) - Mail sender

## Mailer

Object created by [login(argvs)](#login-argvs), used to send mail.

### Methods

| Method | Return type | Summary |
| --- | --- | --- |
| [send(message)](#send-message) | undefined | Sends mail |

## send(message)

Sends mail.

```javascript
mailer.send({
    from: ["Administrator <admin@example.com>"],
    to: ["username@example.com", "UserName <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

### Parameters

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| message | messageArgvs | undefined | true | JavaScript object with the message, see below |

### messageArgvs

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| from | string | undefined | true | Sender email address |
| to | string / string[] | undefined | true | Recipient email address(es) |
| subject | string | undefined | true | Email subject |
| text | string | undefined | true | Plain text body |
| html | string | undefined | false | HTML body |
