# 国考、事业单位职位筛选并导出

脚本基于 `Python3`
首先不太会用 `Excel` 其次 `linux` 里的 `onlyoffice` 有点简陋 自己琢磨了个脚本 希望能助大家上岸。

[Gitee项目地址](https://gitee.com/danyow/py_excel_filter)
[Github项目地址](https://github.com/danyow/py_excel_filter)


# 使用方法

安装 `openpyxl`

# 修改 `main.py` 里面的 `files`

```python

# 修改自己需要的内容

files = [
    {
        'case': '国考.xlsx',  # 查找文件路径
        'skip_rows': [1, 2],  # 需要跳过的行号
        'merge_row': 2,  # 存在 key 但被合并的行好
        'keys_row': 2,  # key 所在行号
        'exports': [
            {
                'file_name': "B-国考.xlsx",  # 导出文件路径
                'filters': {
                    # 以下就是筛选条件
                    '专业': lambda x: have(x, '机械', '工学', '不限'),
                    '学历': lambda x: have(x, '本科', '大专及以上'),
                    '学位': lambda x: x != '硕士',
                    '政治面貌': lambda x: have(x, '不限'),
                    '服务基层项目工作经历': lambda x: have(x, '无限制', '不限'),
                    '基层工作最低年限': lambda x: have(x, '无限制', '不限'),
                    '工作地点': lambda x: have(x, '广东'),
                    '落户地点': lambda x: have(x, '广东'),
                    '备注': lambda x: not_have(x, '女性', '至少具有注册会计师', '大学英语'),
                }
            },
        ]
    },
    {
        'case': '深圳事业单位.xlsx',
        'skip_rows': [1, 2, 3],
        'merge_row': 2,
        'keys_row': 3,
        'exports': [
            {
                'file_name': "B-深圳事业单位.xlsx",
                'filters': {
                    '专业': lambda x: have(x, '机械', '工学', '本科：不限'),
                    '最低专业技术资格': lambda x: must_none(x),
                    '学历': lambda x: have(x, '本科'),
                    '学位': lambda x: have(x, '学士'),
                    '与岗位有关的其它条件': lambda x: not_have(x, '女性', '中共党员', '证', '资格'),
                    '笔试类别': lambda x: not_have(x, '社会人员'),
                }
            },
        ]
    },
]
```

# 运行 && 等待结果