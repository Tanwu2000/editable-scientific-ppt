"""Synthetic regression cases for grouping, clean titles, and print clarity."""
import argparse, copy, json, math, sys
from pathlib import Path
from PIL import Image

HERE=Path(__file__).resolve().parent
sys.path.insert(0,str(HERE.parents[1]/'scripts'))
from native_pptx import build

def text(name,x,y,w,h,label,**style):
    return dict(kind='text',name=name,x=x,y=y,width=w,height=h,text=label,**style)
def path(name,points,**style):
    return dict(kind='shape',name=name,paths=[dict(points=points)],**style)
def panel(left,title):
    right=left+205;top=55;bottom=205
    el=[text('title',left,16,205,24,title,role='title',centerOn=[left,right],size=14),
        path('axes',[[left,top],[left,bottom],[right,bottom]],role='axis'),
        text('x-label',left,231,205,18,'Displacement, δ (µm)',size=10,alignment='center'),
        text('y-label',left-102,122,150,18,'Force, F (nN)',size=10,rotation=-90,alignment='center')]
    for value in [0,0.5,1]:
        x=left+205*value
        el += [path('x-tick',[[x,bottom],[x,bottom+4]],role='axis'),text('tick',x-15,212,30,13,str(value),size=9,alignment='center')]
    for value in [0,1,2]:
        y=bottom-70*value
        el += [path('y-tick',[[left-4,y],[left,y]],role='axis'),text('tick',left-27,y-6,18,13,str(value),size=9,alignment='right')]
    for j,color in enumerate(['#0072B2','#B77811']):
        el.append(path('curve',[[left+205*t/100,bottom-(115+25*j)*(t/100)**(1.25+.2*j)] for t in range(101)],role='curve',stroke=color))
    return el

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--output',type=Path,required=True);args=ap.parse_args()
    args.output.mkdir(parents=True,exist_ok=False)
    # A true mixed group: native paths, text, subgroups and an embedded raster.
    field=Image.new('RGB',(1230,900))
    field.putdata([(round(245-200*math.exp(-((x*900/1230)**2+(y*600/900)**2)/180000)),round(250-120*math.exp(-((x*900/1230)**2+(y*600/900)**2)/180000)),205) for y in range(900) for x in range(1230)])
    field.save(args.output/'field.png',dpi=(432,432))
    a=panel(60,'Editable curves');b=panel(340,'Field and mesh')
    b.insert(1,dict(kind='image',name='synthetic-field',file='field.png',x=340,y=55,width=205,height=150))
    for y in range(80,200,25):b.append(path('mesh',[[330,y],[560,y-20]],role='mesh',clip=[340,55,545,205]))
    mixed=dict(width=600,height=270,elements=[dict(kind='group',name='Panel A',elements=a),dict(kind='group',name='Panel B',elements=b)])
    clean=dict(width=340,height=270,elements=panel(75,'Indentation response'))
    clean['elements'][0].update(x=10,width=100,alignment='left') # must be centered by role
    clean['elements'].append(text('helper',75,40,220,12,'Generated illustration — editable output',size=7,color='#999999',role='helper'))
    baseline=copy.deepcopy(clean);baseline['group']=False
    for e in baseline['elements']:
        e.pop('role',None);e.pop('centerOn',None)
    baseline['elements'][0].update(width=230)
    contrast=copy.deepcopy(clean);contrast['elements']=[e for e in contrast['elements'] if e.get('role')!='helper']
    faint=copy.deepcopy(contrast);faint['group']=False
    for e in faint['elements']:
        if e['kind']=='shape':
            e['lineWidth']=.35
            if e.get('role')=='axis':e['stroke']='#B8B8B8'
        if e.get('role')=='title':e.update(x=75,width=205,alignment='center')
        e.pop('role',None);e.pop('centerOn',None)
    cases={'01_grouping':mixed,'02_layout':clean,'02_layout_before':baseline,'03_clarity':contrast,'03_clarity_before':faint}
    for name,scene in cases.items():
        (args.output/(name+'.json')).write_text(json.dumps(scene,indent=2,ensure_ascii=False),encoding='utf-8')
        print(name,build(scene,args.output,args.output/(name+'.pptx')))

if __name__=='__main__':main()
