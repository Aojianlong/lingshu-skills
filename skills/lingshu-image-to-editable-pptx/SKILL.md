---
name: lingshu-image-to-editable-pptx
description: 当用户上传已有 PNG/JPG 并要求图片转可编辑 PPT 时使用。把流程图、柱状图与表格、照片混排、中英文海报重建为原生可编辑 PPTX，多张图可合成一个文件。模型识图写版面 JSON，固定脚本生成文本框、形状、连接器、原生图表与表格。用于现有图片转换，不替代从大纲创作幻灯片。
---

# 图片转可编辑 PPTX

## 三条硬规则

1. 文字和数字必须是原生对象，不用图片或轮廓冒充，也不把整张原图铺底再叠字。
2. 照片和插画从原图精确裁切，不重新生成；人脸不能被改动。
3. 看不清的数字不猜，要问用户；没做过的检查不说做过。

## 1. 准备环境

运行 `python3 scripts/pptx_tool.py doctor`（Windows 用 `python`）。

- `status` 为 `ready`：继续。
- `missing_dependencies`：在当前工作目录建虚拟环境并安装，之后都用这个环境运行：
  ```sh
  python3 -m venv .pptx-venv
  .pptx-venv/bin/pip install -r <本 Skill 目录>/requirements.txt
  ```
  Windows 的路径是 `.pptx-venv\Scripts\pip`。
- 没有 Python 或装不上依赖：告诉用户缺什么，不返回假文件。

## 2. 看图

取得每张图的实际像素尺寸，逐项读出文字和数字。

多张图按用户给的顺序合成一个 PPTX，每张一页。比例不同的图放不进同一个文件，分开输出并告诉用户。

## 3. 写 scene.json

每张图写一份 scene.json，格式见 [版面协议](references/scene-format.md)。坐标、字号、线宽都用原图像素，元素按背景到前景排序。保留原文、数值、比例和层级，不加装饰。

- 一句或一段用一个文本框，局部粗体用 runs。
- 字体写 `sans`（黑体类）或 `serif`（宋体类），脚本会按系统换成微软雅黑、苹方、宋体。原图明显是英文字体时可写 Arial、Times New Roman。
- 柱状图用原生 chart，柱子上的数字用图表自带标签，带 % 等单位时写 value_format，不另加文本框。
- 表格用原生 table，图标用 shape/line/path，连接线用 start/end 附着形状。
- 折线图、饼图、环图和其他做不成原生图表的图形：从原图裁成 image，并写 `fallback` 说明。页面其余部分照常转换。这些图形的数字不能改，交付时要告诉用户。

## 4. 海报

只有海报需要处理背景，步骤见 [海报处理](references/poster.md)。需要宿主有图像编辑能力，或用户提供无字底图。两者都没有时告诉用户，不要用带字原图代替。

## 5. 生成

```sh
python3 scripts/pptx_tool.py build scene.json --output result.pptx
python3 scripts/pptx_tool.py build page1.json page2.json --output result.pptx
```

输出文件已存在时会自动编号（`result-2.pptx`），实际路径看返回的 `pptx`。只改 scene.json，不改生成脚本。默认按当前系统选字体；用户说明要在另一个系统打开时，加 `--font-target windows` 或 `mac`。

## 6. 看预览

```sh
python3 scripts/render.py result.pptx --output-dir preview
```

Windows 用 PowerPoint，Mac/Linux 用 LibreOffice 和 render-requirements.txt。对照原图和预览，核对文字、数字、版式、溢出和裁切。Mac 的 LibreOffice 可能把苹方粗体显示成宋体，这是预览差异，不用因此改字体。有问题改 scene.json 重新生成，最多改两轮。

返回 `no_renderer` 时，用宿主自带的 PPTX 预览；也没有就直接交付，并说明没有做视觉检查。不要为此让用户安装 LibreOffice。不要用自绘图冒充预览。

海报还要检查背景：运行 `python3 scripts/qa.py edit-check result.pptx --output-dir check`，再渲染其中的 `text-removed-test.pptx`，确认没有旧字和块状修补。

## 7. 交付

给用户 PPTX 和一两句话：哪些能直接改、哪些保留成了图片、哪里和原图有差异。两轮后仍有问题就明确说出问题。
