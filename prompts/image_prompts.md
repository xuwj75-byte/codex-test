# PPT 配图提示词

当前仓库尚未检测到逐字稿，因此暂时无法为每一页生成基于真实内容的配图提示词。

请先将逐字稿放入以下任意位置后运行：

1. `input/transcript.docx`
2. `input/transcript.txt`
3. `input/transcript.md`

运行命令：

```bat
python make_ppt.py
```

脚本会根据实际生成的每一页 PPT，自动补全：页码、页面标题、建议图片类型、中文图片提示词、英文图片提示词、图片风格建议和图片文件建议命名。
