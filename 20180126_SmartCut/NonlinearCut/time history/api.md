## Initial Collection [/initial]

所見即所得，所以沒有 PUT，要更新就直接全部覆蓋。
之所以統一在一個地方，是因為只需要判斷 collection，沒有差太多的邏輯。

### Post Data [POST]

+ Request
    {
        "name": "project",
        "data": [
            {
                "project_id": "T2016-05A",
                "project_name": "FISO"
            },
        ]
    }

## Record Collection [/record/{user_id}/{date}]

核心，完整的 CRUD。

+ Parameters
    + user_id (required, string)
    + date (required, date)

### List User Record [GET]

+ Response 200 (application/json)
    [
        {
            "user_id": "T2016-05A",
            "project_id": "T2016-05A",
            "project_name": "FISO",
            "task": "101 主結構模型",
            "normal_hour": 0,
            "overtime": 1
        },
    ]

### Post User Record [POST]

### Put User Record [PUT]

### Delete User Record [Delete]

## Project Collection [/project/{project_id}/{start_date}/{end_date}]

### Get Project Record [GET]

## Project User Collection [/project-user/{project_id}/{user_id}/{start_date}/{end_date}]

### Get Project User Record [GET]

## User Project Collection [/user-project/{user_id}/{project_id}/{start_date}/{end_date}]

### Get User Project Record [GET]

## User Collection [/user/{user_id}/{date}]

### Get User Record [GET]
