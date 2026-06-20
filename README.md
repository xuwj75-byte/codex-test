# 根据逐字稿生成中文 PPT

这是一个面向 Windows 新手用户的中文 PPT 生成工程。它会优先读取 `input` 文件夹中的逐字稿，并用 `python-pptx` 生成一套适合汇报、讲座或课程使用的简体中文 PPT。

> 当前仓库尚未放入逐字稿，因此不会编造 PPT 内容。请先把逐字稿放入 `input/transcript.docx`、`input/transcript.txt` 或 `input/transcript.md` 后再运行。

## 一、如何放入逐字稿

请在项目文件夹中找到 `input` 文件夹，然后把逐字稿文件放进去。支持以下三种文件名，脚本会按顺序优先读取：

1. `input/transcript.docx`
2. `input/transcript.txt`
3. `input/transcript.md`

如果三个文件都不存在，运行脚本时只会创建 `input/README_请放入逐字稿.md`，提醒你先放入逐字稿，不会生成虚构内容。

## 二、如何安装运行环境

1. 打开 Windows 的“命令提示符”：按 `Win + R`，输入 `cmd`，然后按回车。
2. 进入本项目所在文件夹，例如：

```bat
cd C:\Users\你的用户名\Downloads\codex-test
```

3. 安装依赖：

```bat
python -m pip install -r requirements.txt
```

如果电脑还没有安装 Python，请先到 Python 官网下载安装，并在安装时勾选 **Add python.exe to PATH**。

## 三、如何运行 make_ppt.py

放入逐字稿并安装依赖后，在命令提示符中运行：

```bat
python make_ppt.py
```

脚本会自动生成：

- `output/讲座整理PPT_初版.pptx`
- `output/slides_outline.md`
- `prompts/image_prompts.md`

如果逐字稿内容较短，脚本会尽量生成合理页数；如无法支撑 15—20 页，请补充更多逐字稿内容后重新运行。

## 四、如何查看 output 中的 PPT

运行完成后，打开项目中的 `output` 文件夹，双击：

```text
讲座整理PPT_初版.pptx
```

可以用 Microsoft PowerPoint、WPS 演示或其他兼容 PPTX 的软件查看。

## 五、如何使用 prompts/image_prompts.md 配合 image2 生成图片

`prompts/image_prompts.md` 会为每一页列出：

- 页码
- 页面标题
- 建议图片类型
- 中文图片提示词
- 英文图片提示词
- 图片风格建议
- 图片文件建议命名，例如 `slide_03.png`

使用 image2 时，可以逐页复制对应的中文或英文提示词生成图片。建议保持蓝灰色系、简洁、稳重、适合中文汇报型 PPT，避免夸张、卡通和过度商业广告风格。

## 六、如何把 image2 生成的图片放入 assets 文件夹

1. 在项目文件夹中打开 `assets` 文件夹。
2. 将 image2 生成的图片放入该文件夹。
3. 建议按 `prompts/image_prompts.md` 中的命名保存，例如：

```text
assets\slide_03.png
assets\slide_04.png
assets\slide_05.png
```

这样后续让 Codex 生成带配图的最终版 PPT 时，更容易自动匹配页码和图片。

## 七、后续如何让 Codex 生成带配图的最终版 PPT

当你已经：

1. 放入逐字稿并运行 `python make_ppt.py`；
2. 根据 `prompts/image_prompts.md` 用 image2 生成图片；
3. 将图片保存到 `assets` 文件夹；

就可以让 Codex 继续处理，例如输入：

```text
请读取 assets 文件夹中的 slide_*.png，把图片插入到对应页，生成 output/讲座整理PPT_最终版.pptx。
```

## 八、安全与真实性说明

- 脚本不会上传逐字稿。
- 脚本不会编造逐字稿中没有的关键事实。
- 脚本会删除口水词、重复话和跑题表达，并尽量保留关键思想。
- 第一轮不会直接生成图片，只生成图片提示词。
- 如需更高质量的内容提炼，建议在生成初版后人工核对逐字稿事实。
