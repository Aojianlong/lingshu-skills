# 版面协议 v1

scene.json 描述一张图，坐标全部用原图像素。固定按 96 DPI 转 PPT，模型不要另做字号缩放。

```json
{"version":1,"scene_type":"flowchart","source":"input.png",
 "canvas":{"width":1600,"height":900},"background":"FFFFFF","font":"sans",
 "elements":[
 {"id":"title","type":"text","x":60,"y":45,"w":1400,"h":110,"text":"原图标题","font_px":72,"bold":true,"color":"17354A"},
 {"id":"a","type":"shape","shape":"round_rect","x":70,"y":240,"w":400,"h":240,"radius_px":24,"fill":"E7F4F3","stroke":"008C95","stroke_px":3},
 {"id":"label","type":"text","x":90,"y":320,"w":360,"h":70,"text":"步骤一","font_px":42,"align":"center","valign":"middle"},
 {"id":"arrow","type":"line","x1":470,"y1":360,"x2":650,"y2":360,"color":"008C95","stroke_px":4,"arrow_end":true,"start":{"element":"a","site":3}}
 ]}
```

这是语法示例，不是测试答案。scene_type 取 flowchart、chart_table、photo_layout、poster；一页混合多种内容时用 mixed。每个元素 id 唯一，不得超画布。路径相对 scene.json 或用绝对路径。

font 写 sans（黑体类）或 serif（宋体类）时，脚本按系统换成微软雅黑、苹方、宋体。也可写具体英文字体，如 Arial。

## text

x/y/w/h 加 text 或 paragraphs。font_px 默认32，font 覆盖全局字体，color 为六位 RGB。bold/italic/wrap 为布尔值。align=left/center/right；valign=top/middle/bottom；margin_px 默认0。rotation 可选；vertical=true 表示东亚竖排，需查真实渲染。

```json
{"id":"body","type":"text","x":50,"y":200,"w":600,"h":80,"font_px":32,
 "paragraphs":[{"runs":[{"text":"路线：","bold":true},{"text":"沿江绿道","bold":false}]}]}
```

多行 text 使用换行。paragraphs 每段可覆盖 font_px、align、line_height_px、line_spacing、space_after_px；runs 可覆盖字体/颜色/粗斜体。tracking_px 是字距增量。没有自动缩字号，模型按预览修正框尺寸与字号。

文字填充可用 `"text_fill":{"gradient":["FFF1B8","C18A36"],"angle":90}` 或 `"text_fill":{"texture":"gold.png"}`，仍是原生文字而非标题截图。

## shape、line、path

shape 值：rect、round_rect、ellipse、diamond、triangle；x/y/w/h 必需。fill=RGB、none 或 gradient；stroke/stroke_px 控制轮廓，radius_px 控制圆角，rotation 可选。

line 必需 x1/y1/x2/y2，可加 arrow_start/arrow_end。start/end 为 {"element":"id","site":0} 附着形状；普通矩形常用0上1左2下3右，须实机核对。先形状后连接器更易管理层级。

path 用 points=[[x,y],...]，closed 布尔值及 fill/stroke/stroke_px，生成原生自由形状；不允许把文字轮廓化。

## image

x/y/w/h、asset、role 必需，role=photo/illustration/clean_background。crop_px=[left,top,right,bottom] 为原文件裁切范围，默认拒绝比例不符拉伸。

完整海报底图必须为 clean_background，且与原带字图不同。

做不成原生图表的图形（折线图、饼图等）从原图裁成 image，并写 fallback，生成结果会列出它们：
```json
{"id":"trend","type":"image","role":"illustration","asset":"input.png","crop_px":[60,200,900,800],
 "x":60,"y":200,"w":840,"h":600,"fallback":"折线图保留为图片，数字不能改"}
```

## table

x/y/w/h、rows 二维矩形数组必需，生成原生表格。col_widths、row_heights 可选，各自合计必须为 w/h。header 默认 true；支持 header_fill/header_color/body_fill/color/font_px/border_color/border_px、align。MVP 不支持合并单元格。

## chart

```json
{"id":"chart","type":"chart","chart_type":"column","x":70,"y":300,"w":850,"h":500,
 "categories":["第一期","第二期","第三期"],"series":[{"name":"人数","values":[19,42,76],"color":"0DA597"}],
 "min":0,"max":100,"major_unit":20,"show_values":true,"axis_font_px":24,"label_font_px":26,"gap_width":100}
```

原生图表只支持柱状图，其他图表按上文 fallback 处理。categories 与 values 长度一致，数字必须可读且准确。plot_layout=[x,y,w,h] 可固定绘图区占图表的比例（0..1），不写时脚本按字号留出标签和坐标轴的位置。图表标题另用 text，避免自动布局改变参考图。

柱子上的数字必须用 show_values，不要另写 text，否则改数据时数字不会跟着变。带单位时写 value_format，数据里只放数字：百分比 `"values":[78,65],"value_format":"0\"%\""`，金额 `"value_format":"0\"万\""`。

