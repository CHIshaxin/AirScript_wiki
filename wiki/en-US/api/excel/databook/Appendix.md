---
title: "Appendix"
source: https://airsheet.wps.cn/docs/en-US/api/excel/databook/Appendix.html
updated: 1789700916000
hash: COW1yA_L
---

# Appendix

## Appendix 1: Datasheet field types

| Field type | Type | Create-field shape | Write value | Read value |
| --- | --- | --- | --- | --- |
| Multi-line text | MultiLineText | No extra rules | String | String |
| Date | Date | No extra rules | String, `yyyy/mm/dd` | String |
| Time | Time | No extra rules | String, `hh:mm:ss` | String |
| Number | Number | No extra rules | Number | Number |
| Currency | Currency | No extra rules | Number | Number |
| Percentage | Percentage | No extra rules | Number | Number |
| ID card | ID | No extra rules | String, must match ID rules | String |
| Phone | Phone | No extra rules | String, must match phone rules | String |
| Email | Email | No extra rules | String, must match email rules | String |
| URL | Url | Optional `displayText`: `{"name":"Link","type":"Url","displayText":"Open"}` | String, valid URL | String |
| Checkbox | Checkbox | No extra rules | true / false | Boolean |
| Single select | SingleSelect | At least one option: `{"name": "Option","type": "SingleSelect","items": [{ "value": "item1" }]}` | String, must match an option | String |
| Multi select | MultipleSelect | At least one option: `{"name": "Option","type": "SingleSelect","items": [{ "value": "item1" }, { "value": "item2" }]}` | String array | String array |
| Rating | Rating | `max` between 1 and 5: `{"name": "Rating","type": "Rating","max": 5}` | Number, > 0 and < max | Number |
| Progress | Complete | No extra rules | Number, 0-100 | String |
| Contact | Contact | `multipleContacts` (`<bool>`), `noticeNewContact` (`<bool>`): `{"name": "Contact","type": "Contact","multipleContacts": false,"noticeNewContact": false}` | Not writable | Object |
| Attachment | Attachment | No extra rules | Not writable | Object |
| Link | Link | `linkSheet`, `multipleLinks`: `{"name": "Link","type": "Link","linkSheet": "<sheetId>","multipleLinks": false}` | Array of rows from the linked sheet |  |
| Rich text | Note | No extra rules | Not writable | Object |
| Auto number | AutoNumber | No extra rules | Not writable | Number |
| Created by | CreatedBy | No extra rules | Not writable | Object |
| Created time | CreatedTime | No extra rules | Not writable | String |
| Formula | Formula | No extra rules | Not writable | Depends on formula |
| Lookup | Lookup | No extra rules | Not writable | Same shape as referenced field |

## Appendix 2: Datasheet view types

| View type | Description |
| --- | --- |
| Grid | Grid view |
| Kanban | Kanban view |
| Gallery | Gallery view |
| Form | Form view |
| Gantt | Gantt view |

## Appendix 3: Filter criteria

Filters restrict rows. Two parts: `mode` is how criteria combine; `creteria` holds each rule (`field`, `op`, `values`).

```json
{
  "mode": "AND",
  "criteria": [
    {
      "field": "Name",
      "op": "Intersected",
      "values": ["Datasheet", "12345"]
    },
    {
      "field": "Quantity",
      "op": "Greater",
      "values": ["1"]
    }
  ]
}
```

| Operator | Meaning |
| --- | --- |
| Equals | Equal |
| NotEqu | Not equal |
| Greater | Greater than |
| GreaterEqu | Greater than or equal |
| Less | Less than |
| LessEqu | Less than or equal |
| GreaterEquAndLessEqu | Between (inclusive) |
| LessOrGreater | Between (exclusive) |
| BeginWith | Starts with |
| EndWith | Ends with |
| Contains | Contains |
| NotContains | Does not contain |
| Intersected | One of listed values |
| Empty | Is empty |
| NotEmpty | Is not empty |

Each operator limits how many elements `values` may contain. If the limit is exceeded, that rule is ignored.

`Empty` and `NotEmpty` allow no values; “between” operators allow at most 2; `Intersected` allows up to 65535; others allow at most 1.

Note

`filter` is not a struct: if `field` is missing, `op`/`values` are invalid, `values` is too long, or the rule is otherwise invalid, the whole request fails.

Dynamic date filters use object elements in `values`:

```json
{
  "mode": "AND",
  "criteria": [
    {
      "field": "Date",
      "op": "Equals",
      "values": [
        {
          "dynamicType": "lastMonth",
          "type": "DynamicSimple"
        }
      ]
    }
  ]
}
```

Tip

The example above means “equals last month.”

For dynamic dates, each object in `values` must set `type` to `DynamicSimple`. When `op` is `Equals`, `dynamicType` may be (case-insensitive):

| Value | Meaning |
| --- | --- |
| today | Today |
| yesterday | Yesterday |
| tomorrow | Tomorrow |
| last7Days | Last 7 days |
| last30Days | Last 30 days |
| thisWeek | This week |
| lastWeek | Last week |
| nextWeek | Next week |
| thisMonth | This month |
| lastMonth | Last month |
| nextMonth | Next month |

Tip

When `op` is `greater` or `less`, `dynamicType` may only be yesterday, today, or tomorrow.
