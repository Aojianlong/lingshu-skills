#!/usr/bin/env python3
"""Edit a copy of the PPTX, reread it, and make a text-free copy for checking poster backgrounds."""
import argparse, json, sys
from pathlib import Path
from pptx import Presentation
from pptx.chart.data import CategoryChartData
sys.path.insert(0,str(Path(__file__).resolve().parent))
from pptx_tool import inspect, write_json
def leaves(shapes):
    for s in shapes:
        if s.shape_type==6:yield from leaves(s.shapes)
        else:yield s
def free_dir(path):
    candidate=path;i=2
    while candidate.exists():
        candidate=path.with_name(path.name+"-"+str(i));i+=1
    return candidate
def edit_check(src,out):
    out=free_dir(out)
    out.mkdir(parents=True);p=Presentation(src);edits=[]
    for sn,sl in enumerate(p.slides,1):
        text_done=False
        for s in leaves(sl.shapes):
            if s.has_text_frame and s.text and not text_done:
                before=s.text
                first=next((run for para in s.text_frame.paragraphs for run in para.runs if run.text),None)
                if first is None:first=s.text_frame.paragraphs[0].add_run()
                first.text="EDIT TEST / 编辑验证"
                edits.append({"slide":sn,"shape":s.name,"kind":"text","before":before});text_done=True
            if s.has_table:
                cell=s.table.cell(min(1,len(s.table.rows)-1),min(1,len(s.table.columns)-1));before=cell.text;cell.text="987"
                edits.append({"slide":sn,"shape":s.name,"kind":"table","before":before})
            if s.has_chart:
                chart=s.chart;cd=CategoryChartData();cd.categories=[c.label for c in chart.plots[0].categories]
                for i,ser in enumerate(chart.series):
                    vals=list(ser.values)
                    if i==0:before=vals[0];vals[0]=float(vals[0])+1
                    cd.add_series(ser.name,vals)
                chart.replace_data(cd);edits.append({"slide":sn,"shape":s.name,"kind":"chart","before":before,"after":before+1})
    target=out/"edited-test.pptx";p.save(target);audit=inspect(target)
    for e in edits:
        it=next(i for i in audit["slides"][e["slide"]-1] if i["id"]==e["shape"])
        if e["kind"]=="text":assert "EDIT TEST" in it["text"]
        elif e["kind"]=="table":assert any("987" in row for row in it["table"])
        else:assert it["chart"]["series"][0]["values"][0]==e["after"]
    p=Presentation(src)
    for sl in p.slides:
        for s in list(leaves(sl.shapes)):
            if s.has_text_frame and s.text:
                el=s._element;el.getparent().remove(el)
    no_text=out/"text-removed-test.pptx";p.save(no_text)
    return {"status":"passed","check":"file edit and reread, not PowerPoint editing","output_dir":str(out),
            "edited_pptx":target.name,"edits":edits,"text_free_copy":no_text.name}
def main():
    a=argparse.ArgumentParser(description=__doc__);sub=a.add_subparsers(dest="cmd",required=True)
    e=sub.add_parser("edit-check");e.add_argument("pptx");e.add_argument("--output-dir",required=True)
    args=a.parse_args()
    try:
        print(json.dumps(edit_check(Path(args.pptx),Path(args.output_dir)),ensure_ascii=False,indent=2));return 0
    except Exception as exc:print(json.dumps({"status":"failed","error":str(exc)},ensure_ascii=False),file=sys.stderr);return 1
if __name__=="__main__":sys.exit(main())

