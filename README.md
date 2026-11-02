# 灵枢 Skills

灵枢（lingshu）出品的 AI Agent Skills，主要给 Codex 用。

| Skill | 用途 |
| --- | --- |
| [lingshu-image-to-editable-pptx](skills/lingshu-image-to-editable-pptx) | 把图片转成原生可编辑的 PPT |

## lingshu-image-to-editable-pptx：图片转可编辑 PPT

把流程图、柱状图和表格、照片混排、中英文海报等图片，重建成原生可编辑的 PPTX。文字可以直接改，大小和颜色可以调，柱状图可以右键「编辑数据」。多张图可以合成一个 PPT，每张一页。

### 安装

**方法一：把链接发给 Codex**

在 Codex 里发送：

```text
帮我安装这个 Skill：https://github.com/Aojianlong/lingshu-skills/tree/main/skills/lingshu-image-to-editable-pptx
```

**方法二：下载后拖进 Codex**

1. 在本页点 **Code → Download ZIP**，解压。
2. 打开 Codex，把 `skills/lingshu-image-to-editable-pptx` 整个文件夹拖进对话框，发送：

```text
帮我安装这个 Skill
```

装好后开一个新对话再用。

### 使用

把图片拖进或粘贴到 Codex 对话框，然后发送：

```text
$lingshu-image-to-editable-pptx 转为可编辑 PPT
```

也可以在对话框输入 `/` 或 `$`，从列表里选「lingshu-image-to-editable-pptx（灵枢·图片转可编辑PPT）」，再输入「转为可编辑 PPT」。

生成后打开 PPT 检查：点一下文字直接改；柱状图右键选「编辑数据」，改完柱子会跟着变。如果新数字超过坐标轴上限，再右键坐标轴把最大值调大。

### 能转什么

| 图片类型 | 转完能改什么 |
| --- | --- |
| 流程图 | 文字、方框、箭头 |
| 柱状图和表格 | 文字、数据、表格内容 |
| 照片混排 | 文字和形状；照片是单独的图片，可以移动和替换 |
| 中英文海报 | 文字；背景是一张去掉文字的图片。需要 Codex 能修图，或者你提供无字底图 |

折线图、饼图目前不能转成原生图表，会保留成图片，页面其他内容照常转换。转换是看图重建，字体和细节可能和原图略有差异。

运行需要 Python 3。第一次使用时，Codex 会自动在当前文件夹建虚拟环境并安装依赖（python-pptx、Pillow）。

## 交流和更新

遇到问题或想要后续更新，微信扫码加我：

<img src="assets/wechat-qr.png" width="260" alt="微信二维码">
