## Initial Collection [/initial]

所見即所得，所以沒有 PUT，要更新就直接全部覆蓋。
之所以統一在一個地方，是因為只需要判斷 collection，沒有差太多的邏輯。

### Post Data [POST]

+ Request (application/json)

    {
        "name": "project",
        "data": [
            {
                "project_id": "T2016-05A",
                "project_name": "FISO"
            },
            {
                "project_id": "H2018-04A",
                "project_name": "廣慈A"
            },
        ]
    }

+ Response 201 (application/json)

    {
        "name": "project",
        "data": [
            {
                "_id": "123",
                "project_id": "T2016-05A",
                "project_name": "FISO"
            },
            {
                "_id": "456",
                "project_id": "H2018-04A",
                "project_name": "廣慈A"
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
            "_id": "234",
            "user_id": "EI 201705",
            "user_name": "Willy",
            "date": "2018/12/1",
            "project_id": "T2016-05A",
            "project_name": "FISO",
            "task": "101 主結構模型",
            "normal_hour": 0,
            "overtime": 1
        },
        {
            "_id": "567",
            "user_id": "EI 201705",
            "user_name": "Willy",
            "date": "2018/12/1",
            "project_id": "T2016-05A",
            "project_name": "FISO",
            "task": "102 開挖分析",
            "normal_hour": 2,
            "overtime": 0
        },
    ]

### Post User Record [POST]

+ Request (application/json)

    [
        {
            "user_id": "EI 201705",
            "user_name": "Willy",
            "date": "2018/12/1",
            "project_id": "T2016-05A",
            "project_name": "FISO",
            "task": "101 主結構模型",
            "normal_hour": 0,
            "overtime": 1
        },
        {
            "user_id": "EI 201705",
            "user_name": "Willy",
            "date": "2018/12/1",
            "project_id": "T2016-05A",
            "project_name": "FISO",
            "task": "102 開挖分析",
            "normal_hour": 2,
            "overtime": 0
        },
    ]

+ Response 201 (application/json)

    [
        {
            "_id": "234",
            "user_id": "EI 201705",
            "user_name": "Willy",
            "date": "2018/12/1",
            "project_id": "T2016-05A",
            "project_name": "FISO",
            "task": "101 主結構模型",
            "normal_hour": 0,
            "overtime": 1
        },
        {
            "_id": "567",
            "user_id": "EI 201705",
            "user_name": "Willy",
            "date": "2018/12/1",
            "project_id": "T2016-05A",
            "project_name": "FISO",
            "task": "102 開挖分析",
            "normal_hour": 2,
            "overtime": 0
        },
    ]

### Put User Record [PUT]

+ Request (application/json)

    [
        {
            "_id": "234",
            "normal_hour": 1,
        },
    ]

+ Response 201 (application/json)

    [
        {
            "_id": "234",
            "user_id": "EI 201705",
            "user_name": "Willy",
            "date": "2018/12/1",
            "project_id": "T2016-05A",
            "project_name": "FISO",
            "task": "101 主結構模型",
            "normal_hour": 1,
            "overtime": 1
        },
    ]

### Delete User Record [Delete]

+ Request (application/json)

    [
        {
            "_id": "234",
        },
    ]

+ Response 204

## Project Collection [/project/{project_id}/{start_date}/{end_date}]

這裡的重點是想把邏輯放在哪裡，想放在後端處理完。

### Get Project Record [GET]

+ Response 200 (application/msword)

## Project User Collection [/project-user/{project_id}/{user_id}/{start_date}/{end_date}]

### Get Project User Record [GET]

+ Response 200 (application/msword)

## User Project Collection [/user-project/{user_id}/{project_id}/{start_date}/{end_date}]

### Get User Project Record [GET]

+ Response 200 (application/msword)

## User Collection [/user/{user_id}/{date}]

### Get User Record [GET]

+ Response 200 (application/msword)

