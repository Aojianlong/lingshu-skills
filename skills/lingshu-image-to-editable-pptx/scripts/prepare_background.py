#!/usr/bin/env python3
"""Blend a host image-editor's clean background only into declared text regions."""
import argparse, json
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--source",required=True);ap.add_argument("--edited",required=True)
    ap.add_argument("--regions",required=True,help="JSON list of [left,top,right,bottom] in source pixels")
    ap.add_argument("--output",required=True);ap.add_argument("--feather",type=float,default=5)
    a=ap.parse_args();out=Path(a.output)
    if out.suffix.lower()!=".png":ap.error("Clean background must be lossless PNG")
    base=out;i=2
    while any(f.exists() for f in [out,out.with_suffix(".mask.png")]):
        out=base.with_name(base.stem+"-"+str(i)+base.suffix);i+=1
    source=Image.open(a.source).convert("RGB");edited=Image.open(a.edited).convert("RGB")
    if abs((source.width/source.height)/(edited.width/edited.height)-1)>.01:ap.error("Image editor changed framing/aspect ratio; obtain a matching edit")
    edited=edited.resize(source.size,Image.Resampling.LANCZOS)
    boxes=json.loads(Path(a.regions).read_text(encoding="utf-8-sig"))
    if not boxes:ap.error("At least one text region required")
    # The caller inspects the EDITED asset first. No inpainting, OCR or AI is hidden here.
    mask=Image.new("L",source.size,0);d=ImageDraw.Draw(mask)
    for b in boxes:
        if len(b)!=4 or not (0<=b[0]<b[2]<=source.width and 0<=b[1]<b[3]<=source.height):ap.error("Invalid region")
        d.rectangle(b,fill=255)
    if a.feather<0 or a.feather>30:ap.error("Feather must be 0..30 pixels")
    if a.feather:mask=mask.filter(ImageFilter.GaussianBlur(a.feather))
    result=Image.composite(edited,source,mask)
    out.parent.mkdir(parents=True,exist_ok=True);result.save(out)
    mask.save(out.with_suffix(".mask.png"))
    print(json.dumps({"status":"written","output":str(out),"regions":boxes,"note":"Pixels outside the regions match the source. Inspect the result before using it."},ensure_ascii=False,indent=2))
if __name__=="__main__":main()

