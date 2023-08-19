# 项目更新日志

本文件记录当前项目的更新情况，历史更改记录。
项目基础的开发工具需求：vscode，git，安装 vscode 插件 Microsoft Python 和 Pylance

* 步骤一建立属于项目本地的 Python 环境

```shell
python3 -m venv .myvenv

# 在非 vscode 终端环境中需要通过命令进入本地 Python 环境
source .myvenv/bin/activate

```

* 步骤二建立项目的测试环境

VS Code 单元测试功能默认是关闭的，需要在 settings.json 里面添加一行：
"python.testing.pytestEnabled": true
以上配置使用的是第三方的 pytest 模块。
你也可以使用标准库里的 unittest："python.testing.unittestEnabled": true
或者另外一个第三方模块 nose："python.testing.nosetestsEnabled": true
但是这三个不同的测试模块同时只能启用一个。

## 更新记录

* 2022/4/02 创建这个项目，支持本地的 Python 虚拟环境。

## 参考资源

Poetry vs Pip 作为包管理工具
[Starting New Python Project in VSCode](https://zhauniarovich.com/post/2020/2020-04-starting-new-python-project/)

## To-Do 列表
