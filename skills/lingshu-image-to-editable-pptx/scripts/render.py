#!/usr/bin/env python3
"""Render an actual PPTX with LibreOffice or Windows PowerPoint; never a simulated preview."""
import argparse, base64, json, os, platform, shutil, subprocess, sys, tempfile, re
from xml.sax.saxutils import escape
from pathlib import Path
def next_free_dir(path):
    # A new folder for every run, so old previews never mix with new ones.
    candidate=path;i=2
    while candidate.exists() and any(candidate.iterdir()):
        candidate=path.with_name(path.name+"-"+str(i));i+=1
    return candidate
def unavailable(status, message):
    print(json.dumps({"status":status,"message":message},ensure_ascii=False,indent=2))
    sys.exit(0)
def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument("pptx");p.add_argument("--output-dir",required=True)
    p.add_argument("--backend",choices=["auto","libreoffice","powerpoint"],default="auto")
    p.add_argument("--font-dir",action="append",default=[],help="Extra font directory for this render only; does not install fonts")
    a=p.parse_args();src=Path(a.pptx).resolve();out=Path(a.output_dir).resolve()
    if not src.is_file():p.error("PPTX not found")
    out=next_free_dir(out)
    out.mkdir(parents=True,exist_ok=True)
    backend=a.backend
    if backend=="auto":
        backend="libreoffice"
        if platform.system()=="Windows":
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,"PowerPoint.Application\\CLSID"):
                    backend="powerpoint"
            except OSError:
                pass
    images=[];missing_texts=[]
    if backend=="powerpoint":
        if platform.system()!="Windows":p.error("PowerPoint renderer requires Windows and a signed-in desktop")
        # Local-only COM. Does not install Office, activate it, or change execution policy.
        def sq(s):return str(s).replace("'","''")
        ps=r"""$ErrorActionPreference='Stop'
$app=New-Object -ComObject PowerPoint.Application
$app.Visible=-1
$p=$app.Presentations.Open('__SRC__',-1,0,-1)
try{
 for($i=1;$i -le $p.Slides.Count;$i++){
  $w=[int][Math]::Round($p.PageSetup.SlideWidth/72*96)
  $h=[int][Math]::Round($p.PageSetup.SlideHeight/72*96)
  $p.Slides.Item($i).Export((Join-Path '__OUT__' ("slide-"+$i+".png")),'PNG',$w,$h)
 }
}finally{$p.Close()}
""".replace("__SRC__",sq(src)).replace("__OUT__",sq(out))
        encoded=base64.b64encode(ps.encode("utf-16-le")).decode("ascii")
        subprocess.run(["powershell.exe","-NoProfile","-NonInteractive","-STA","-EncodedCommand",encoded],check=True,timeout=180)
        images=sorted(out.glob("slide-*.png"))
    else:
        office=shutil.which("soffice") or shutil.which("libreoffice")
        mac=Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        if not office and mac.exists():office=str(mac)
        if not office:unavailable("no_renderer","No local renderer. Use the host's PPTX preview if it has one; otherwise deliver the PPTX and say it was not visually checked. Do not install LibreOffice unless the user asks.")
        try:import pymupdf as fitz
        except ImportError:unavailable("missing_pymupdf","LibreOffice is available; install render-requirements.txt to get PNG previews.")
        with tempfile.TemporaryDirectory(prefix="editable-pptx-") as td:
            profile=Path(td,"profile").as_uri()
            env=os.environ.copy()
            if platform.system()=="Darwin" or a.font_dir:
                dirs=[Path(x).resolve() for x in a.font_dir]
                dirs += [Path.home()/"Library/Fonts",Path("/Library/Fonts"),Path("/System/Library/Fonts"),Path("/System/Library/Fonts/Supplemental")] if platform.system()=="Darwin" else [Path("/usr/share/fonts"),Path("/usr/local/share/fonts"),Path.home()/".local/share/fonts"]
                config=Path(td,"fonts.conf")
                config.write_text('<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd"><fontconfig>'+''.join('<dir>'+escape(str(d))+'</dir>' for d in dirs if d.exists())+'<cachedir>'+escape(str(Path(td,"font-cache")))+'</cachedir></fontconfig>',encoding="utf-8")
                env["FONTCONFIG_FILE"]=str(config)
                env["FONTCONFIG_PATH"]=td
            r=subprocess.run([office,"-env:UserInstallation="+profile,"--headless","--convert-to","pdf","--outdir",str(out),str(src)],capture_output=True,text=True,timeout=180,env=env)
            pdf=out/(src.stem+".pdf")
            if r.returncode or not pdf.exists():p.error("LibreOffice export failed: "+r.stdout+r.stderr)
            doc=fitz.open(pdf)
            for i,page in enumerate(doc,1):
                f=out/("slide-"+str(i)+".png");page.get_pixmap(matrix=fitz.Matrix(96/72,96/72),alpha=False).save(f);images.append(f)
            text=''.join(page.get_text() for page in doc)
            from pptx import Presentation
            normalize=lambda s: re.sub(r"\s+","",s)
            actual=normalize(text)
            for slide in Presentation(src).slides:
                for shape in slide.shapes:
                    candidates=[shape.text] if shape.has_text_frame else []
                    if shape.has_table:candidates += [cell.text for row in shape.table.rows for cell in row.cells]
                    for expected in candidates:
                        if len(normalize(expected))>1 and normalize(expected) not in actual:missing_texts.append(expected)
            doc.close()
    if not images:p.error("Renderer produced no PNG")
    manifest={"status":"rendered","renderer":backend,"output_dir":str(out),
              "images":[f.name for f in images],"text_coverage_warnings":missing_texts}
    print(json.dumps(manifest,ensure_ascii=False,indent=2))
if __name__=="__main__":main()

