# export.py

Export excel to Python/Json/Lua

参考 https://github.com/yanghuan/proton

## Excel 文件格式

### 1. 通用格式

Excel 格式是这样的：

| 索引            | 名称     | 坐骑            | 性别            | 身高            |
|---------------|--------|---------------|---------------|---------------|
| int           | str    | int           | int           | int           |
| Id            | Name   | MountId       | Sex           | Height        |
| server/client | client | server/client | server/client | server/client |
| 1             | 奥丁     | 10001         | 1             | 181           |
| 2             | 托尔     | 10002         | 1             | 186           |
| 3             | 希芙     |               | 0             | 172           |
| '             |        |               |               |               |
| 6             | 希芙6    |               | 0             | 172           |


- 前四行为表头，从第五行开始为数据
- 第一行为描述
- 第二行定义数据类型，支持整数(int),字符串(str),整数数组(arrint),字符串数组(arrstr)
- 第三行定义属性的 KEY
- 第四行设定是否导出标识，比如 `server/client` 标识 `server` 和 `client` 都导出

```
python3 ../export.py -f ./hero.xlsx -s 英雄 -t ./hero_data.json
```

导出的数据格式是这样的：

```json
{
    "1": {
        "Height": 181,
        "MountId": 10001,
        "Name": "奥丁",
        "Sex": 1
    },
    "2": {
        "Height": 186,
        "MountId": 10002,
        "Name": "托尔",
        "Sex": 1
    },
    "3": {
        "Height": 172,
        "Name": "希芙",
        "Sex": 0
    },
    "6": {
        "Height": 172,
        "Name": "希芙6",
        "Sex": 0
    }
}
```


### 2. 全局格式

Excel 格式是这样的：

| name          | value          | type   | out           | description |
|---------------|----------------|--------|---------------|-------------|
| NameLimit     | 7              | int    | client/server | 名字字数限制      |
| HitCorrection | 1              | int    | server        | 命中修正系数      |
| InitItem      | 1001,1002,1003 | arrint | server        | 初始任务获得道具    |
| '             |                |        |               |             |
| InitMission   | M101,M102      | arrstr | server        | 初始任务        |

- 第一行为表头，`name` 为属性 KEY 列，`value` 列为数据列
- 第二行开始为数据

```
python3 ../export.py -f ./hero.xlsx -s 全局参数 -t ./global_data.json -k global
```

导出的数据格式是这样的：

```json
{
    "HitCorrection": 1,
    "InitItem": [
        1001,
        1002,
        1003
    ],
    "InitMission": [
        "M101",
        "M102"
    ],
    "NameLimit": 7
}
```

## 使用方法

```
python3 export.py -h
```

参考 `test/test.py`

推荐编写自动化脚本，读取 `conf.ini` 文件，实现批量导表。可以集成版本管理工具 Svn/Git 到自动化脚本。同时也推荐集成到服务器指令模块。

## 导出数据给 C# 使用

如果是 Unity 里使用，推荐使用 C# 编写工具解析 json 格式，然后再用 `SaveAssets` 序列化。使用的时候直接 Load。

