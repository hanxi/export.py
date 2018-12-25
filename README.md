# export.py

Export excel to Python/Json/Lua

## Excel 文件格式

### 1. 通用模式

### 2. 全局模式

## 使用方法

参考 `test/test.py`

推荐编写自动化脚本，读取 `conf.ini` 文件，实现批量导表。可以集成版本管理工具 Svn/Git 到自动化脚本。

## 导出数据给 C# 使用

如果是 Unity 里使用，推荐使用 C# 编写工具解析 json 格式，然后再用 `SaveAssets` 序列化。使用的时候直接 Load。

