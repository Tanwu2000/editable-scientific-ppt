"""Package actual PowerPoint renders as a reproducible comparison gallery."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from matplotlib.font_manager import findfont, FontProperties
from zipfile import ZipFile, ZIP_DEFLATED
import argparse, copy, json, shutil, sys
from datetime import date
from xml.etree import ElementTree as ET

HERE=Path(__file__).resolve().parent
sys.path.insert(0,str(HERE.parents[1]/'scripts'))
from inspect_pptx import inspect

def public_deck(source, target):
    # Strip Office profile metadata only; all slide content stays byte-identical.
    with ZipFile(source) as src, ZipFile(target,'w',ZIP_DEFLATED) as dst:
        for entry in src.infolist():
            data=src.read(entry.filename)
            if entry.filename=='docProps/core.xml':
                root=ET.fromstring(data)
                for el in root:
                    if el.tag.rsplit('}',1)[-1] in ('creator','lastModifiedBy'):el.text=''
                data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
            dst.writestr(copy.copy(entry),data)
    with ZipFile(source) as src, ZipFile(target) as dst:
        for name in src.namelist():
            if name.startswith('ppt/slides/'):
                assert src.read(name)==dst.read(name)

def main():
    ap=argparse.ArgumentParser();ap.add_argument('artifacts',type=Path)
    ap.add_argument('--tested-date',default=date.today().isoformat())
    ap.add_argument('--output',type=Path,default=HERE/'results');args=ap.parse_args()
    if args.output.exists():raise FileExistsError('Use a fresh publication directory')
    args.output.mkdir(parents=True)
    font=ImageFont.truetype(findfont(FontProperties(family='Arial')),27)
    report=json.loads((args.artifacts/'powerpoint-report.json').read_text(encoding='utf-8-sig'))
    for row in report['cases']:
        src=args.artifacts/row['case'];dst=args.output/row['case'];dst.mkdir()
        for name in ['source.svg','source.png','svg-inserted.png','svg-converted.png','native.png','native-edited.png','native_geometry_check.json']:
            shutil.copy2(src/name,dst/name)
        for name in ['native.pptx','svg-converted.pptx','native-edited.pptx']:
            public_deck(src/name,dst/name)
        before=Image.open(src/'svg-converted.png').convert('RGB')
        after=Image.open(src/'native.png').convert('RGB')
        assert before.size==after.size
        width=950;height=round(width*before.height/before.width)
        canvas=Image.new('RGB',(1980,height+100),'white');draw=ImageDraw.Draw(canvas)
        for x,label,pic in [(20,'SVG → Shapes',before),(1010,'Native editable PPTX',after)]:
            draw.text((x+width/2,27),label,fill='#111111',font=font,anchor='mt')
            canvas.paste(pic.resize((width,height),Image.Resampling.LANCZOS),(x,80))
        draw.line((990,80,990,height+80),fill='#DDDDDD',width=1)
        canvas.save(dst/'comparison.png',optimize=True)
        row['structure']={}
        for name in ['native','svg-converted']:
            result=inspect(dst/(name+'.pptx'));result.pop('texts')
            row['structure'][name]=result
    report['testedDate']=args.tested_date
    report['data']='Synthetic deterministic fixtures; not experimental or finite-element results.'
    report['comparison']='Same source SVG; actual PowerPoint slide exports; no retouching inside plots.'
    (args.output/'report.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')

if __name__=='__main__':main()
