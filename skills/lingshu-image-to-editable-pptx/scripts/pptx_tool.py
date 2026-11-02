#!/usr/bin/env python3
"""Deterministic scene JSON -> native PPTX. No model API or Office dependency."""
from __future__ import annotations
import argparse, hashlib, io, json, math, os, platform, re, shutil, sys, zipfile
from pathlib import Path
from collections import Counter
try:
    from PIL import Image
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
    from pptx.chart.data import CategoryChartData
    from pptx.oxml.xmlchemy import OxmlElement
    from lxml import etree
except ImportError as exc:
    print(json.dumps({"status":"missing_dependencies","missing":exc.name,"python":sys.version.split()[0],
                      "fix":"python -m pip install -r requirements.txt"},ensure_ascii=False))
    sys.exit(3)

VERSION = "0.2.0"
PX = 9525  # physical pixel at 96 DPI; NEVER scale font a second time.
NS = {"p":"http://schemas.openxmlformats.org/presentationml/2006/main",
      "a":"http://schemas.openxmlformats.org/drawingml/2006/main",
      "c":"http://schemas.openxmlformats.org/drawingml/2006/chart",
      "r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
KINDS = {"text", "shape", "line", "path", "image", "table", "chart"}
SHAPES = {"rect":MSO_SHAPE.RECTANGLE, "round_rect":MSO_SHAPE.ROUNDED_RECTANGLE,
          "ellipse":MSO_SHAPE.OVAL, "diamond":MSO_SHAPE.DIAMOND,
          "triangle":MSO_SHAPE.ISOSCELES_TRIANGLE}
# Theme-independent font roles, resolved for the system that will open the file.
FONT_ROLES = {"windows":{"sans":"Microsoft YaHei","serif":"SimSun"},
              "mac":{"sans":"PingFang SC","serif":"Songti SC"}}
# Numeric scene fields measured in canvas pixels; crop_px stays in source-file pixels.
SCALED = {"x","y","w","h","x1","y1","x2","y2","font_px","stroke_px","border_px","tracking_px","radius_px",
          "line_height_px","space_after_px","margin_px","axis_font_px","label_font_px"}
class SceneError(ValueError): pass

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def next_free(path):
    """Keep existing files: result.pptx -> result-2.pptx."""
    path=Path(path);i=2;candidate=path
    while candidate.exists():
        candidate=path.with_name(path.stem+"-"+str(i)+path.suffix);i+=1
    return candidate
def write_json(path, data):
    path=next_free(path)
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    return path
def font_target(value="auto"):
    if value!="auto":return value
    return "mac" if platform.system()=="Darwin" else "windows"
def apply_fonts(node, roles):
    if isinstance(node,dict):
        def pick(v):
            if str(v).startswith("Noto Sans CJK"):v="sans"
            elif str(v).startswith("Noto Serif CJK"):v="serif"
            return roles.get(v,v)
        return {k:(pick(v) if k=="font" else apply_fonts(v,roles)) for k,v in node.items()}
    if isinstance(node,list):return [apply_fonts(v,roles) for v in node]
    return node
def scale_scene(node, f):
    if isinstance(node,list):return [scale_scene(v,f) for v in node]
    if not isinstance(node,dict):return node
    out={}
    for k,v in node.items():
        if k in SCALED and isinstance(v,(int,float)) and not isinstance(v,bool):out[k]=v*f
        elif k=="points":out[k]=[[x*f,y*f] for x,y in v]
        elif k in {"col_widths","row_heights"}:out[k]=[x*f for x in v]
        elif k in {"crop_px","cleaning"}:out[k]=v
        else:out[k]=scale_scene(v,f)
    return out
def color(v):
    v=str(v).lstrip("#")
    if not re.fullmatch("[0-9a-fA-F]{6}",v): raise SceneError("Expected six-digit RGB: "+v)
    return RGBColor.from_string(v.upper())
def num(x, label):
    if isinstance(x,bool) or not isinstance(x,(int,float)) or not math.isfinite(x):
        raise SceneError("Expected finite number: "+label)
    return float(x)
def px(x): return Emu(round(float(x)*PX))
def resolve(root, p):
    if not isinstance(p,str) or not p: raise SceneError("Missing asset path")
    if re.match(r"^[a-z]+://",p,re.I): raise SceneError("Assets must be local files, not network URLs")
    q=Path(p).expanduser()
    return q.resolve() if q.is_absolute() else (root/q).resolve()
def xml(parent, tag, **attrs):
    e=OxmlElement(tag)
    for k,v in attrs.items(): e.set(k,str(v))
    parent.append(e)
    return e
def remove_fill(parent):
    for c in list(parent):
        if etree.QName(c).localname in {"solidFill","noFill","gradFill","blipFill","pattFill","grpFill"}:
            parent.remove(c)
def ordered_fill(parent, child):
    parent.remove(child)
    followers={"ln","effectLst","effectDag","highlight","uLnTx","uLn","uFillTx","uFill","latin","ea","cs","sym","hlinkClick","hlinkMouseOver","rtl","extLst"}
    index=next((i for i,c in enumerate(parent) if etree.QName(c).localname in followers),len(parent))
    parent.insert(index,child)
    return child
def set_fill(parent, spec, slide=None, root=None):
    remove_fill(parent)
    if spec is None or spec=="none": ordered_fill(parent,xml(parent,"a:noFill")); return
    if isinstance(spec,str): xml(ordered_fill(parent,xml(parent,"a:solidFill")),"a:srgbClr",val=str(color(spec))); return
    if "gradient" in spec:
        grad=ordered_fill(parent,xml(parent,"a:gradFill",rotWithShape="1"))
        gs=xml(grad,"a:gsLst")
        stops=spec["gradient"]
        if len(stops)<2:raise SceneError("Gradient requires two or more color stops")
        for i,c in enumerate(stops):
            item=xml(gs,"a:gs",pos=round(i*100000/(len(stops)-1)))
            xml(item,"a:srgbClr",val=str(color(c)))
        xml(grad,"a:lin",ang=round(spec.get("angle",90)*60000),scaled="1")
    elif "texture" in spec:
        if slide is None: raise SceneError("Texture needs a slide")
        part,rid=slide.part.get_or_add_image_part(str(resolve(root,spec["texture"])))
        fill=ordered_fill(parent,xml(parent,"a:blipFill",dpi="96",rotWithShape="1"))
        b=xml(fill,"a:blip"); b.set("{"+NS["r"]+"}embed",rid)
        xml(fill,"a:tile",tx="0",ty="0",sx="100000",sy="100000",algn="tl")
    else: raise SceneError("Fill must be RGB, gradient, texture or none")
def style_font(font, style, default):
    font.name=style.get("font",default)
    for tag in ["ea","cs"]:
        child=font._rPr.find("{"+NS["a"]+"}"+tag)
        if child is None:child=xml(font._rPr,"a:"+tag)
        child.set("typeface",style.get("font",default))
    font.size=Pt(style.get("font_px",32)*0.75)
    font.bold=bool(style.get("bold",False)); font.italic=bool(style.get("italic",False))
    font.color.rgb=color(style.get("color","162F42"))
def style_run(run, style, default, slide, root):
    style_font(run.font,style,default)
    rpr=run._r.get_or_add_rPr()
    for tag in ["a:ea","a:cs"]:
        child=rpr.find("{"+NS["a"]+"}"+tag.split(":")[1])
        if child is None: child=xml(rpr,tag)
        child.set("typeface",style.get("font",default))
    rpr.set("lang",style.get("lang","zh-CN"))
    rpr.set("spc",str(round(style.get("tracking_px",0)*75)))
    if "text_fill" in style: set_fill(rpr,style["text_fill"],slide,root)
def text_frame(tf, e, default, slide, root):
    tf.clear(); tf.word_wrap=bool(e.get("wrap",False)); tf.auto_size=MSO_AUTO_SIZE.NONE
    margin=e.get("margin_px",0)
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=px(margin)
    tf.vertical_anchor={"top":MSO_ANCHOR.TOP,"middle":MSO_ANCHOR.MIDDLE,"bottom":MSO_ANCHOR.BOTTOM}[e.get("valign","top")]
    # Whole paragraphs remain editable; use runs only for mixed styles.
    paras=e.get("paragraphs")
    if paras is None: paras=[{"text":s} for s in e.get("text","").split("\n")]
    for i,pd in enumerate(paras):
        if isinstance(pd,str): pd={"text":pd}
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        st={**e,**pd}
        p.alignment={"left":PP_ALIGN.LEFT,"center":PP_ALIGN.CENTER,"right":PP_ALIGN.RIGHT}[st.get("align","left")]
        p.space_before=Pt(0); p.space_after=Pt(st.get("space_after_px",0)*.75)
        p.line_spacing=Pt(st["line_height_px"]*.75) if "line_height_px" in st else st.get("line_spacing",1.0)
        for rd in pd.get("runs",[{"text":pd.get("text","")}]):
            run=p.add_run();run.text=str(rd["text"]);style_run(run,{**st,**rd},default,slide,root)
        ppr=p._p.get_or_add_pPr()
        xml(ppr,"a:buNone")
    if e.get("vertical"): tf._txBody.bodyPr.set("vert","eaVert")

def union_area(rects):
    xs=sorted({v for r in rects for v in (r[0],r[2])});area=0
    for left,right in zip(xs,xs[1:]):
        ys=sorted((r[1],r[3]) for r in rects if r[0]<right and r[2]>left)
        length=0;end=None
        for lo,hi in ys:
            if end is None or lo>end:length+=hi-lo;end=hi
            elif hi>end:length+=hi-end;end=hi
        area+=(right-left)*length
    return area

def validate(scene, root):
    if scene.get("version")!=1: raise SceneError("Scene version must be 1")
    cv=scene.get("canvas",{})
    w=num(cv.get("width"),"canvas.width");h=num(cv.get("height"),"canvas.height")
    if min(w,h)<32 or max(w,h)>5376: raise SceneError("Canvas outside 32..5376 px (PowerPoint size limit)")
    if scene.get("scene_type") not in {"flowchart","chart_table","photo_layout","poster","mixed"}:
        raise SceneError("scene_type must be flowchart, chart_table, photo_layout, poster or mixed")
    reference=resolve(root,scene.get("source"))
    with Image.open(reference) as im:
        if im.size!=(round(w),round(h)): raise SceneError("Canvas must match source image's exact pixel size")
    els=scene.get("elements")
    if not isinstance(els,list) or not els: raise SceneError("Nonempty elements required")
    ids=set();warnings=[];counts=Counter()
    for e in els:
        eid=e.get("id")
        if not isinstance(eid,str) or not eid or eid in ids: raise SceneError("Unique nonempty element id required")
        ids.add(eid);kind=e.get("type");counts[kind]+=1
        if kind not in KINDS: raise SceneError("Unsupported native element: "+str(kind))
        if kind=="line":
            for k in ("x1","y1","x2","y2"): num(e.get(k),eid+"."+k)
            points=[[e["x1"],e["y1"]],[e["x2"],e["y2"]]]
        elif kind=="path":
            points=e.get("points",[])
            if len(points)<2: raise SceneError(eid+": path needs 2+ points")
            for pt in points:
                if len(pt)!=2: raise SceneError("Path point must be [x,y]")
                for v in pt:num(v,eid)
        else:
            for k in ("x","y","w","h"):num(e.get(k),eid+"."+k)
            if min(e["w"],e["h"])<=0:raise SceneError(eid+": nonpositive dimensions")
            points=[[e["x"],e["y"]],[e["x"]+e["w"],e["y"]+e["h"]]]
        if any(x<-.5 or y<-.5 or x>w+.5 or y>h+.5 for x,y in points):
            raise SceneError(eid+": element extends outside source canvas")
        if kind=="text":
            if not e.get("text") and not e.get("paragraphs"):raise SceneError(eid+": empty text")
            size=num(e.get("font_px",32),eid+".font_px")
            if size<=0:raise SceneError(eid+": invalid font size")
            if e.get("font_px",32)>e["h"]*1.6:warnings.append(eid+": font size exceeds frame; inspect native render")
        if kind=="shape" and e.get("shape","rect") not in SHAPES:raise SceneError("Unsupported shape preset")
        if kind=="image":
            path=resolve(root,e.get("asset"))
            if not path.is_file():raise SceneError("Missing image: "+str(path))
            role=e.get("role")
            if role not in {"photo","illustration","clean_background"}:raise SceneError(eid+": image role required")
            with Image.open(path) as im:
                crop=e.get("crop_px")
                if crop and (len(crop)!=4 or crop[0]<0 or crop[1]<0 or crop[2]>im.width or crop[3]>im.height or crop[2]<=crop[0] or crop[3]<=crop[1]):
                    raise SceneError(eid+": invalid pixel crop [left,top,right,bottom]")
                iw,ih=(crop[2]-crop[0],crop[3]-crop[1]) if crop else im.size
            if abs((iw/ih)/(e["w"]/e["h"])-1)>.02 and not e.get("allow_stretch",False):
                raise SceneError(eid+": asset aspect ratio differs; use a precise crop, not stretching")
            full=e["w"]*e["h"]/(w*h)>.90
            if full and role!="clean_background":raise SceneError("Full-slide bitmap permitted only as reviewed clean background")
            if full and sha(path)==sha(reference):raise SceneError("Original whole-slide image cannot be a clean background")
            if role=="clean_background":
                if sha(path)==sha(reference):raise SceneError("Unchanged source is not a clean background")
                warnings.append(eid+": inspect no-text native render for residual letters and damaged texture")
        if kind=="table":
            rows=e.get("rows",[])
            if not rows or not rows[0] or any(len(r)!=len(rows[0]) for r in rows):raise SceneError(eid+": rectangular table required")
            if "col_widths" in e and (len(e["col_widths"])!=len(rows[0]) or abs(sum(e["col_widths"])-e["w"])>1):raise SceneError(eid+": column widths must sum to table width")
        if kind=="chart":
            if e.get("chart_type","column")!="column":raise SceneError("Only native column charts are supported; keep other charts as an image element with a fallback note")
            cats=e.get("categories",[]);series=e.get("series",[])
            if not cats or not series:raise SceneError("Chart needs categories and series")
            for s in series:
                if len(s.get("values",[]))!=len(cats):raise SceneError("Chart category/value count mismatch")
                for v in s["values"]:num(v,eid+".values")
    for e in els:
        for key in ("start","end"):
            a=e.get(key)
            if isinstance(a,dict) and a.get("element") not in ids:raise SceneError("Connector target does not exist")
    raster_rects=[[e["x"],e["y"],e["x"]+e["w"],e["y"]+e["h"]] for e in els if e["type"]=="image" and e.get("role")!="clean_background"]
    if union_area(raster_rects)>w*h*.90:raise SceneError("Raster tiles collectively cover the slide; do not disguise the original as editable elements")
    if counts["text"]==0:raise SceneError("At least one native text element required")
    if scene["scene_type"]=="chart_table" and not (counts["chart"] or counts["table"] or any(e.get("fallback") for e in els)):
        raise SceneError("chart_table needs a native chart or table; unsupported charts need an image element with a fallback note")
    if scene["scene_type"]=="poster" and not any(e.get("role")=="clean_background" for e in els):
        raise SceneError("Poster requires a separately cleaned background")
    return warnings

def no_effects(shape):
    # Explicitly remove theme references; an empty effectLst alone is ignored by some renderers.
    for theme_style in list(shape._element.findall("{"+NS["p"]+"}style")):
        shape._element.remove(theme_style)
    sp=shape._element.find("{"+NS["p"]+"}spPr")
    if sp is not None:
        for c in list(sp):
            if etree.QName(c).localname in {"effectLst","effectDag"}:sp.remove(c)
        xml(sp,"a:effectLst")
def shape_style(s,e,slide,root):
    no_effects(s);set_fill(s._element.spPr,e.get("fill"),slide,root)
    stroke=e.get("stroke")
    if stroke:
        s.line.color.rgb=color(stroke);s.line.width=Pt(e.get("stroke_px",2)*.75)
    else:s.line.fill.background()
    if e.get("rotation"):s.rotation=e["rotation"]
def table_shape(slide,e,default,root):
    rows=e["rows"];s=slide.shapes.add_table(len(rows),len(rows[0]),px(e["x"]),px(e["y"]),px(e["w"]),px(e["h"]))
    tab=s.table
    widths=e.get("col_widths",[e["w"]/len(rows[0])]*len(rows[0]))
    heights=e.get("row_heights",[e["h"]/len(rows)]*len(rows))
    if len(heights)!=len(rows) or abs(sum(heights)-e["h"])>1:raise SceneError("row_heights must sum to table height")
    for j,v in enumerate(widths):tab.columns[j].width=px(v)
    for i,v in enumerate(heights):tab.rows[i].height=px(v)
    for i,row in enumerate(rows):
        for j,value in enumerate(row):
            cell=tab.cell(i,j);head=i==0 and e.get("header",True)
            style={**e,"text":str(value),"font_px":e.get("font_px",28),"align":e.get("align","center"),"valign":"middle",
                   "color":e.get("header_color","FFFFFF") if head else e.get("color","233448"),"bold":head}
            cell.fill.solid();cell.fill.fore_color.rgb=color(e.get("header_fill","087C8C") if head else e.get("body_fill","FFFFFF"))
            text_frame(cell.text_frame,style,default,slide,root)
            cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            tc=cell._tc.get_or_add_tcPr()
            for side in ["L","R","T","B"]:
                tag="{"+NS["a"]+"}ln"+side
                for old in list(tc.findall(tag)):tc.remove(old)
                ln=xml(tc,"a:ln"+side,w=round(e.get("border_px",1)*PX))
                tc.remove(ln)
                before=next((k for k,node in enumerate(tc) if etree.QName(node).localname in {"cell3D","noFill","solidFill","gradFill","blipFill","pattFill","grpFill","headers","extLst"}),len(tc))
                tc.insert(before,ln)
                xml(xml(ln,"a:solidFill"),"a:srgbClr",val=str(color(e.get("border_color","D5DCE1"))))
    return s
def chart_shape(slide,e,default):
    data=CategoryChartData();data.categories=e["categories"]
    for ss in e["series"]:data.add_series(ss.get("name","Series"),ss["values"])
    frame=slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,px(e["x"]),px(e["y"]),px(e["w"]),px(e["h"]),data)
    chart=frame.chart;chart.has_legend=e.get("legend",False);chart.has_title=False
    style_font(chart.font,e,default)
    ca=chart.category_axis;va=chart.value_axis
    for axis in [ca,va]:
        style_font(axis.tick_labels.font,{**e,"font_px":e.get("axis_font_px",24),"bold":False},default)
        axis.format.line.color.rgb=color(e.get("axis_color","9BA3AA"))
        axis.format.line.width=Pt(.6)
    va.minimum_scale=e.get("min",0)
    if "max" in e:va.maximum_scale=e["max"]
    if "major_unit" in e:va.major_unit=e["major_unit"]
    va.has_major_gridlines=True;va.major_gridlines.format.line.color.rgb=color(e.get("grid_color","E2E6EA"))
    va.major_gridlines.format.line.width=Pt(.5)
    plot=chart.plots[0];plot.gap_width=int(e.get("gap_width",100))
    plot.has_data_labels=e.get("show_values",True)
    if plot.has_data_labels:
        plot.data_labels.position=XL_LABEL_POSITION.OUTSIDE_END
        style_font(plot.data_labels.font,{**e,"font_px":e.get("label_font_px",26),"bold":True},default)
    for ser,info in zip(chart.series,e["series"]):
        ser.format.fill.solid();ser.format.fill.fore_color.rgb=color(info.get("color","0DA597"))
        ser.format.line.fill.background()
    # Reserve room for data labels and axes; automatic layouts differ between PowerPoint and WPS.
    layout_box=e.get("plot_layout")
    if not layout_box:
        left=e.get("axis_font_px",24)*2.6/e["w"];bottom=e.get("axis_font_px",24)*2.2/e["h"]
        top=e.get("label_font_px",26)*1.8/e["h"] if plot.has_data_labels else .03
        layout_box=[left,top,1-left-.02,1-top-bottom]
    # PowerPoint presses a label onto a bar near the axis top. Extend the axis past its last
    # tick and grow the plot upward by the same pixels, so bars and ticks stay where they were.
    if plot.has_data_labels and "max" in e and "major_unit" in e:
        layout_box=list(layout_box)
        axis_min=e.get("min",0);span=e["max"]-axis_min
        top_value=max(v for ss in e["series"] for v in ss["values"])
        plot_h=layout_box[3]*e["h"];per_value=plot_h/span
        extra_px=e.get("label_font_px",26)*1.6-(e["max"]-top_value)*per_value
        if extra_px>0:
            extra_px=min(extra_px,layout_box[1]*e["h"])
            va.maximum_scale=round(e["max"]+extra_px/per_value,6)
            layout_box[1]-=extra_px/e["h"];layout_box[3]+=extra_px/e["h"]
            # Ticks stay whole numbers; otherwise PowerPoint reserves width for the long axis maximum.
            if "value_format" not in e and float(e["major_unit"]).is_integer() and float(axis_min).is_integer():
                va.tick_labels.number_format="0";va.tick_labels.number_format_is_linked=False
    if e.get("value_format"):
        va.tick_labels.number_format=e["value_format"];va.tick_labels.number_format_is_linked=False
        if plot.has_data_labels:
            plot.data_labels.number_format=e["value_format"];plot.data_labels.number_format_is_linked=False
    if layout_box:
        pa=chart._chartSpace.find(".//{"+NS["c"]+"}plotArea")
        old=pa.find("{"+NS["c"]+"}layout")
        if old is not None:pa.remove(old)
        layout=OxmlElement("c:layout");pa.insert(0,layout);ml=xml(layout,"c:manualLayout")
        xml(ml,"c:layoutTarget",val="inner")
        for key in ["x","y"]:xml(ml,"c:"+key+"Mode",val="edge")
        for key,val in zip(["x","y","w","h"],layout_box):xml(ml,"c:"+key,val=val)
    return frame

def add_slide(pres, scene, root):
    slide=pres.slides.add_slide(pres.slide_layouts[6])
    slide.background.fill.solid();slide.background.fill.fore_color.rgb=color(scene.get("background","FFFFFF"))
    default=scene.get("font","sans")
    shapes={};pending=[]
    for e in scene["elements"]:
        kind=e["type"];s=None
        if kind=="text":
            s=slide.shapes.add_textbox(px(e["x"]),px(e["y"]),px(e["w"]),px(e["h"]))
            text_frame(s.text_frame,e,default,slide,root);no_effects(s)
            if e.get("rotation"):s.rotation=e["rotation"]
        elif kind=="shape":
            s=slide.shapes.add_shape(SHAPES[e.get("shape","rect")],px(e["x"]),px(e["y"]),px(e["w"]),px(e["h"]))
            shape_style(s,e,slide,root)
            if e.get("shape")=="round_rect" and "radius_px" in e:
                s.adjustments[0]=min(.5,max(0,e["radius_px"]/min(e["w"],e["h"])))
        elif kind=="line":
            s=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,px(e["x1"]),px(e["y1"]),px(e["x2"]),px(e["y2"]))
            s.line.color.rgb=color(e.get("color","008C95"));s.line.width=Pt(e.get("stroke_px",3)*.75)
            if e.get("arrow_end"):xml(s._element.spPr.get_or_add_ln(),"a:tailEnd",type="triangle",w="sm",len="sm")
            if e.get("arrow_start"):xml(s._element.spPr.get_or_add_ln(),"a:headEnd",type="triangle",w="sm",len="sm")
            pending.append((s,e))
        elif kind=="path":
            points=e["points"];builder=slide.shapes.build_freeform(points[0][0],points[0][1],scale=PX)
            builder.add_line_segments(points[1:],close=e.get("closed",False));s=builder.convert_to_shape()
            shape_style(s,e,slide,root)
        elif kind=="image":
            path=resolve(root,e["asset"])
            with Image.open(path) as im:
                if e.get("crop_px"):im=im.crop(tuple(e["crop_px"]))
                buf=io.BytesIO();im.save(buf,format="PNG");buf.seek(0)
            s=slide.shapes.add_picture(buf,px(e["x"]),px(e["y"]),px(e["w"]),px(e["h"]))
        elif kind=="table":s=table_shape(slide,e,default,root)
        elif kind=="chart":s=chart_shape(slide,e,default)
        no_effects(s)
        s.name=e["id"];shapes[e["id"]]=s
    for s,e in pending:
        for end,method in [("start",s.begin_connect),("end",s.end_connect)]:
            anchor=e.get(end)
            if anchor:method(shapes[anchor["element"]],int(anchor.get("site",0)))

def build(scene_files, output, target="auto"):
    """One scene becomes one slide; all slides share the first scene's size."""
    if isinstance(scene_files,(str,Path)):scene_files=[scene_files]
    out=Path(output).resolve()
    if out.suffix.lower()!=".pptx":raise SceneError("PPT output must use .pptx")
    target=font_target(target);roles=FONT_ROLES[target]
    pres=Presentation();warnings=[];fallbacks=[];width=None
    # Validate everything first so a bad page leaves no partial deck behind.
    loaded=[]
    for name in scene_files:
        scene_file=Path(name).resolve();root=scene_file.parent
        scene=json.loads(scene_file.read_text(encoding="utf-8-sig"))
        warnings+=validate(scene,root);loaded.append((scene_file,root,scene))
    for scene_file,root,scene in loaded:
        w,h=scene["canvas"]["width"],scene["canvas"]["height"]
        if width is None:
            width,height=w,h
            pres.slide_width=px(w);pres.slide_height=px(h)
        elif abs((w/h)/(width/height)-1)>.01:
            raise SceneError(scene_file.name+": aspect ratio differs from the first page; build it as a separate PPTX")
        scene=apply_fonts(scene,roles)
        if "font" not in scene:scene["font"]=roles["sans"]
        if w!=width:scene=scale_scene(scene,width/w)
        add_slide(pres,scene,root)
        fallbacks+=[{"scene":scene_file.name,"id":e["id"],"note":e["fallback"]} for e in scene["elements"] if e.get("fallback")]
    out=next_free(out)
    out.parent.mkdir(parents=True,exist_ok=True)
    pres.save(out)
    return {"version":VERSION,"pptx":str(out),"slides":len(loaded),"font_target":target,
            "status":"needs_visual_review","warnings":warnings,"fallbacks":fallbacks,"audit":inspect(out)}

def inspect(file):
    p=Presentation(file);slides=[]
    for sl in p.slides:
        items=[]
        def walk(shapes):
            for s in shapes:
                if s.shape_type==6:yield from walk(s.shapes)
                else:yield s
        for s in walk(sl.shapes):
            item={"id":s.name,"type":int(s.shape_type),"x":s.left/PX,"y":s.top/PX,"w":s.width/PX,"h":s.height/PX}
            if s.has_text_frame:item["text"]=s.text_frame.text
            if s.has_table:item["table"]=[[c.text for c in row.cells] for row in s.table.rows]
            if s.has_chart:
                item["chart"]={"categories":[c.label for c in s.chart.plots[0].categories],
                               "series":[{"name":z.name,"values":list(z.values)} for z in s.chart.series]}
            items.append(item)
        slides.append(items)
    return {"slides":slides,"width_px":p.slide_width/PX,"height_px":p.slide_height/PX,
            "counts":{"text":sum(bool(i.get("text")) for sl in slides for i in sl),
                      "pictures":sum(i["type"]==13 for sl in slides for i in sl),
                      "tables":sum("table" in i for sl in slides for i in sl),
                      "charts":sum("chart" in i for sl in slides for i in sl)}}

def main():
    ap=argparse.ArgumentParser(description=__doc__);sub=ap.add_subparsers(dest="cmd",required=True)
    sub.add_parser("doctor")
    b=sub.add_parser("build");b.add_argument("scene",nargs="+");b.add_argument("--output",required=True)
    b.add_argument("--font-target",choices=["auto","windows","mac"],default="auto")
    i=sub.add_parser("inspect");i.add_argument("pptx");i.add_argument("--output")
    a=ap.parse_args()
    try:
        if a.cmd=="doctor":
            import pptx
            office=shutil.which("soffice") or shutil.which("libreoffice") or (Path("/Applications/LibreOffice.app").exists() or None)
            r={"status":"ready","python":sys.version.split()[0],"platform":platform.system(),"architecture":platform.machine(),
               "python_pptx":pptx.__version__,"font_target":font_target(),"fonts":FONT_ROLES[font_target()],
               "preview":"powerpoint_or_host" if platform.system()=="Windows" else ("libreoffice" if office else "host_only")}
        elif a.cmd=="build":r=build(a.scene,a.output,a.font_target)
        else:
            r=inspect(a.pptx)
            if a.output:write_json(a.output,r)
        print(json.dumps(r,ensure_ascii=False,indent=2))
    except Exception as e:
        print(json.dumps({"status":"failed","error":str(e)},ensure_ascii=False),file=sys.stderr);return 1
    return 0
if __name__=="__main__":sys.exit(main())

