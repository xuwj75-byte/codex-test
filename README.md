# Windows 本地文件整理工具

这是一个面向 Windows 新手用户的本地文件整理工具。它只整理你指定的某一个文件夹，不会扫描全盘，不会删除文件，也不会修改文件内容。

## 这个工具会做什么

工具会按照“文件类型 + 最后修改年份”生成整理结构，例如：

```text
目标文件夹
├─ Word文档
│  └─ 2026
├─ PDF文件
│  └─ 2025
├─ Excel表格
├─ PPT演示
├─ 图片
├─ 视频
├─ 音频
├─ 压缩包
├─ 文本文件
└─ 其他文件
```

默认情况下，工具只会生成整理计划，不会真的移动文件。只有你在命令中添加 `--execute`（Python 版本）或 `-Execute`（PowerShell 版本），才会正式移动或复制文件。

## 安全承诺

- 不会删除文件。
- 不会修改文件内容。
- 不会默认移动文件。
- 只会扫描你通过 `--source` 指定的文件夹。
- 如果目标位置已经有同名文件，不会覆盖旧文件，而是自动改名，例如 `文件名_001.pdf`。
- 预演模式会生成 `organize_plan.csv`，方便你先检查。
- 每次运行会写入 `organize_log.txt` 日志文件。


## 未安装 Python？优先使用 PowerShell 版本

如果你的电脑没有安装 Python，建议优先使用 `organize_files.ps1`。PowerShell 是 Windows 自带的命令行工具，Windows 10 和 Windows 11 通常都可以直接使用，所以不需要先安装 Python，也更适合不熟悉编程的新手用户。

> 下面的 PowerShell 版本和 Python 版本一样安全：默认只预演，只整理你指定的 `Source` 文件夹，不扫描全盘，不删除文件，不修改文件内容，也不会覆盖同名文件。

### 如何打开命令提示符 CMD

1. 按键盘上的 `Win + R`。
2. 输入 `cmd`。
3. 按回车，打开黑色窗口“命令提示符”。

如果脚本放在桌面，可以先进入桌面：

```bat
cd Desktop
```

如果脚本放在下载文件夹，可以进入下载文件夹：

```bat
cd Downloads
```

也可以进入你实际保存脚本的文件夹，例如：

```bat
cd C:\Users\你的用户名\Downloads
```

### 如何运行 PowerShell 脚本

在 CMD 里运行 PowerShell 脚本时，可以使用下面这种格式：

```bat
powershell -ExecutionPolicy Bypass -File .\organize_files.ps1 -Source "C:\Users\你的用户名\Downloads" -Target "D:\整理后的文件"
```

请把示例路径换成你自己的路径：

- `-Source`：你要整理的原文件夹。
- `-Target`：你希望整理后文件放到哪里。

### PowerShell 版本：先预演，不移动文件

PowerShell 版本默认就是预演模式 dry-run。下面命令只会生成整理计划，不会移动、复制或删除文件：

```bat
powershell -ExecutionPolicy Bypass -File .\organize_files.ps1 -Source "C:\Users\你的用户名\Downloads" -Target "D:\整理后的文件"
```

运行后会生成两个文件：

- `organize_plan.csv`：整理计划表。
- `organize_log.txt`：运行日志。

### 如何查看 organize_plan.csv

1. 找到 `organize_plan.csv`。
2. 双击打开，通常会用 Excel 打开。
3. 重点检查这些列：
   - 原文件路径：文件现在在哪里。
   - 目标文件路径：文件将要放到哪里。
   - 文件类型：会进入哪个分类文件夹。
   - 文件大小：文件大小，单位是字节。
   - 最后修改时间：用于决定年份文件夹。
   - 是否会重名：如果显示“是”，脚本会自动在文件名后加编号，避免覆盖。

如果计划不符合你的预期，请不要执行正式命令。可以修改 `-Source` 或 `-Target` 后重新预演。

### 确认后使用 -Copy -Execute 复制整理

如果确认 `organize_plan.csv` 没问题，并且希望原文件夹保留一份、目标文件夹再复制一份，请运行：

```bat
powershell -ExecutionPolicy Bypass -File .\organize_files.ps1 -Source "C:\Users\你的用户名\Downloads" -Target "D:\整理后的文件" -Copy -Execute
```

这里的意思是：

- `-Execute`：正式执行。没有这个参数时永远只是预演。
- `-Copy`：复制文件，不从原文件夹移走文件。

如果你真的想移动文件，可以只加 `-Execute`，不加 `-Copy`：

```bat
powershell -ExecutionPolicy Bypass -File .\organize_files.ps1 -Source "C:\Users\你的用户名\Downloads" -Target "D:\整理后的文件" -Execute
```

移动后，文件会从原文件夹转移到目标文件夹中。新手更建议先用 `-Copy -Execute`。

### 使用 PowerShell 版本时如何避免误操作

- 第一次一定只运行预演命令，不要加 `-Execute`。
- 一定要打开 `organize_plan.csv` 仔细检查。
- 新手建议正式执行时优先使用 `-Copy -Execute`，先复制整理，确认无误后再手动处理原文件。
- 建议先用一个小文件夹测试，例如新建一个“测试整理”文件夹，放几个文件进去试运行。
- `-Source` 不要写成 `C:\`、`D:\` 这类磁盘根目录，避免整理范围过大。
- `-Target` 建议使用一个新的空文件夹，例如 `D:\整理后的文件`。
- 不要把 `-Source` 写错。脚本只会整理你指定的 `Source` 文件夹，不会扫描全盘。

## Python 版本使用方法

## 第一步：安装 Python

1. 打开浏览器，访问 Python 官网：https://www.python.org/downloads/
2. 点击下载 Windows 版本的 Python。
3. 双击安装包。
4. 在安装界面最重要的一步：勾选 **Add python.exe to PATH**。
5. 点击 **Install Now**。
6. 安装完成后，点击 **Close**。

### 检查 Python 是否安装成功

1. 按键盘上的 `Win + R`。
2. 输入 `cmd`，按回车，打开“命令提示符”。
3. 输入下面命令并按回车：

```bat
python --version
```

如果看到类似 `Python 3.x.x` 的版本号，说明安装成功。

## 第二步：打开命令提示符

1. 按 `Win + R`。
2. 输入 `cmd`。
3. 按回车。

如果你的 `organize_files.py` 放在桌面，可以先进入桌面：

```bat
cd Desktop
```

如果脚本放在其他文件夹，请使用 `cd` 命令进入脚本所在位置。例如：

```bat
cd C:\Users\你的用户名\Downloads
```

## 第三步：先预演，不移动文件

预演模式是默认模式。下面命令只会生成整理计划，不会移动文件：

```bat
python organize_files.py --source "C:\Users\你的用户名\Downloads" --target "D:\整理后的文件"
```

请把示例路径换成你自己的路径：

- `--source`：你要整理的原文件夹。
- `--target`：你希望整理后文件放到哪里。

运行后会生成两个文件：

- `organize_plan.csv`：整理计划表。
- `organize_log.txt`：运行日志。

## 第四步：查看整理计划

1. 找到 `organize_plan.csv`。
2. 双击打开，通常会用 Excel 打开。
3. 检查每一行：
   - 原文件路径：文件现在在哪里。
   - 目标文件路径：文件将要放到哪里。
   - 文件类型：会进入哪个分类文件夹。
   - 文件大小：文件大小，单位是字节。
   - 最后修改时间：用于决定年份文件夹。
   - 是否会重名：如果显示“是”，工具会自动加编号，避免覆盖。

如果计划不符合你的预期，请不要执行正式命令，可以修改 `--source` 或 `--target` 后重新预演。

## 第五步：确认后正式执行移动

确认 `organize_plan.csv` 没问题后，再添加 `--execute` 正式移动文件：

```bat
python organize_files.py --source "C:\Users\你的用户名\Downloads" --target "D:\整理后的文件" --execute
```

注意：移动后，文件会从原文件夹转移到目标文件夹中。

## 如果你只想复制，不想移动

如果你希望原文件夹保留一份，目标文件夹再复制一份，请同时添加 `--execute --copy`：

```bat
python organize_files.py --source "C:\Users\你的用户名\Downloads" --target "D:\整理后的文件" --execute --copy
```

这样会复制文件，不会从原文件夹移走。

## 如何避免误操作

- 第一次一定只运行预演命令，不要加 `--execute`。
- 一定要打开 `organize_plan.csv` 仔细检查。
- 建议先用一个小文件夹测试，例如新建一个“测试整理”文件夹，放几个文件进去试运行。
- `--source` 不要写成 `C:\`、`D:\` 这类磁盘根目录，避免整理范围过大。
- `--target` 建议使用一个新的空文件夹，例如 `D:\整理后的文件`。
- 如果不确定，请使用 `--copy`，这样原文件夹还会保留文件。

## 常见问题

### 会不会删除我的文件？

不会。这个工具没有删除文件的功能。

### 会不会覆盖同名文件？

不会。如果目标位置已有同名文件，会自动加编号，例如 `照片_001.jpg`。

### 为什么没有移动文件？

因为默认是预演模式。必须添加 `--execute` 才会正式执行。

### 日志在哪里？

默认在当前命令所在文件夹生成 `organize_log.txt`。
