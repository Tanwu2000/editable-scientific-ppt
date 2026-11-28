"""Compose comparison panels from unmodified actual PowerPoint renders."""
import argparse,json,shutil,sys,struct,zlib
from pathlib import Path
from datetime import date
from PIL import Image,ImageDraw,ImageFont,ImageChops,ImageStat
from matplotlib.font_manager import findfont,FontProperties

HERE=Path(__file__).resolve().parent
sys.path.insert(0,str(HERE.parent/'gallery'))
from publish_results import public_deck
sys.path.insert(0,str(HERE.parents[1]/'scripts'))
from inspect_pptx import inspect

def save_tiff(im,path,dpi=600):
    # Baseline RGB TIFF with Adobe Deflate: avoids a platform libtiff dependency.
    rgb=im.convert('RGB');data=zlib.compress(rgb.tobytes(),6)
    extra=8+2+13*12+4
    entries=[(256,4,1,rgb.width),(257,4,1,rgb.height),(258,3,3,extra),
        (259,3,1,8),(262,3,1,2),(273,4,1,extra+22),(277,3,1,3),
        (278,4,1,rgb.height),(279,4,1,len(data)),(282,5,1,extra+6),
        (283,5,1,extra+14),(284,3,1,1),(296,3,1,2)]
    blob=b'II'+struct.pack('<HIH',42,8,len(entries))
    blob+=b''.join(struct.pack('<HHII',*e) for e in sorted(entries))+struct.pack('<I',0)
    blob+=struct.pack('<HHHIIII',8,8,8,dpi,1,dpi,1)+data
    path.write_bytes(blob)
    assert zlib.decompress(blob[extra+22:])==rgb.tobytes()

def main():
    ap=argparse.ArgumentParser();ap.add_argument('artifacts',type=Path);ap.add_argument('--output',type=Path,default=HERE/'results');ap.add_argument('--tested-date',default=date.today().isoformat());args=ap.parse_args()
    args.output.mkdir(parents=True,exist_ok=False)
    font=ImageFont.truetype(findfont(FontProperties(family='Arial')),27)
    results={}
    for name in ['01_grouping','02_layout','03_clarity','02_layout_before','03_clarity_before']:
        qa=args.artifacts/(name+'-qa-final')
        report=json.loads((qa/'powerpoint-check.json').read_text(encoding='utf-8-sig'))
        scene=json.loads((args.artifacts/(name+'.json')).read_text(encoding='utf-8'))
        im=Image.open(qa/'preview.png').convert('RGB')
        assert im.size==(round(scene['width']/72*600),round(scene['height']/72*600))
        diff=ImageChops.difference(im,Image.open(qa/'regrouped.png').convert('RGB'))
        report['RegroupMeanPixelError']=ImageStat.Stat(diff).mean
        assert max(report['RegroupMeanPixelError'])<.05,'Regroup layout changed'
        report['Structure']=inspect(args.artifacts/(name+'.pptx'));report['Structure'].pop('texts')
        results[name]=report
        public_deck(args.artifacts/(name+'.pptx'),args.output/(name+'.pptx'))
        shutil.copy2(args.artifacts/(name+'.json'),args.output/(name+'.json'))
        im.save(args.output/(name+'.png'),dpi=(600,600),optimize=True)
        if not name.endswith('_before'):
            save_tiff(im,args.output/(name+'.tiff'))
            public_deck(qa/'editability-test-copy.pptx',args.output/(name+'-edited.pptx'))
            shutil.copy2(qa/'edited.png',args.output/(name+'-edited.png'))
            shutil.copy2(qa/'regrouped.png',args.output/(name+'-regrouped.png'))
    shutil.copy2(args.artifacts/'field.png',args.output/'field.png')
    for name in ['02_layout','03_clarity']:
        a=Image.open(args.output/(name+'_before.png'));b=Image.open(args.output/(name+'.png'))
        w=850;h=round(w*a.height/a.width)
        canvas=Image.new('RGB',(1760,h+85),'white');draw=ImageDraw.Draw(canvas)
        for x,label,im in [(20,'Issue reproduction',a),(890,'v0.1.2',b)]:
            draw.text((x+w/2,20),label,font=font,fill='black',anchor='mt')
            canvas.paste(im.resize((w,h),Image.Resampling.LANCZOS),(x,70))
        canvas.save(args.output/(name+'-comparison.png'),optimize=True)
    im=Image.open(args.output/'01_grouping.png');im.thumbnail((1500,900));im.save(args.output/'01_grouping-preview.png')
    (args.output/'report.json').write_text(json.dumps({'date':args.tested_date,'cases':results},indent=2)+'\n',encoding='utf-8')

if __name__=='__main__':main()
