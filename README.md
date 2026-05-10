# HermesTrayHelper

HermesTrayHelper 是一个 Windows 托盘启动器，用 PowerShell 5.1 和 Windows Forms 封装常用 Hermes 命令，并桥接到 WSL2 Ubuntu 中运行的 Hermes。

## 功能

- 托盘菜单启动 Hermes、继续最近会话、浏览 CLI 会话。
- 动态加载最近 CLI 会话，并支持启动、重命名、删除会话。
- Gateway 子菜单支持启动、重启、状态查看、关闭，并按当前状态隐藏不适用操作。
- Gateway 状态和短操作使用 Windows 对话框显示，不弹出多余命令行窗口。

## 使用

双击运行：

```bat
Start-HermesHelper.bat
```

调试模式：

```bat
Start-HermesHelper.bat debug
```

自检：

```bat
Start-HermesHelper.bat selftest
```

## 前提

- Windows PowerShell 5.1
- WSL2 Ubuntu
- WSL 内可以直接运行 `hermes`

