"""Bounded native PowerPoint scene writer. Python 3.10+, standard library only."""
from __future__ import annotations
import argparse
import copy
import json
import math
from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
for prefix, uri in NS.items(): ET.register_namespace(prefix, uri)

def node(tag, attrs=None, parent=None, text=None):
    prefix, name = tag.split(':')
    e = ET.Element('{%s}%s' % (NS[prefix], name), {k: str(v) for k, v in (attrs or {}).items()})
    if parent is not None: parent.append(e)
    if text is not None: e.text = text
    return e

def xml(e): return ET.tostring(e, encoding='utf-8', xml_declaration=True)
def emu(v): return round(float(v) * 12700)
def number(v):
    if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v):
        raise ValueError('Expected a finite number')
    return v

def fields(obj, allowed):
    if not isinstance(obj, dict): raise ValueError('Expected an object')
    extra = set(obj) - set(allowed.split())
    if extra: raise ValueError('Unsupported fields: ' + ', '.join(sorted(extra)))

def color(v):
    if v != 'none' and not re.fullmatch(r'#[0-9a-fA-F]{6}', v):
        raise ValueError('Expected #RRGGBB or none: ' + str(v))
    return v

def segment(a, b, box):
    if box is None: return [list(a), list(b)]
    x, y = a; dx, dy = b[0]-x, b[1]-y; lo, hi = 0., 1.
    for p, q in [(-dx,x-box[0]), (dx,box[2]-x), (-dy,y-box[1]), (dy,box[3]-y)]:
        if abs(p) < 1e-14:
            if q < 0: return None
        elif p < 0: lo = max(lo, q/p)
        else: hi = min(hi, q/p)
        if lo > hi: return None
    return [[x+lo*dx,y+lo*dy],[x+hi*dx,y+hi*dy]]

def convex(pts):
    signs = []
    for i, b in enumerate(pts):
        a, c = pts[i-1], pts[(i+1)%len(pts)]
        cross = (b[0]-a[0])*(c[1]-b[1]) - (b[1]-a[1])*(c[0]-b[0])
        if abs(cross) > 1e-10: signs.append(cross > 0)
    return bool(signs) and all(s == signs[0] for s in signs)

def polygon(pts, box):
    if box is None: return pts
    for axis, limit, sign in [(0,box[0],1),(0,box[2],-1),(1,box[1],1),(1,box[3],-1)]:
        if not pts: break
        out=[]; prev=pts[-1]; pin=(prev[axis]-limit)*sign >= 0
        for cur in pts:
            cin=(cur[axis]-limit)*sign >= 0
            if cin != pin:
                t=(limit-prev[axis])/(cur[axis]-prev[axis])
                out.append([prev[j]+t*(cur[j]-prev[j]) for j in (0,1)])
            if cin: out.append(cur)
            prev,pin=cur,cin
        pts=out
    return pts

def clipped_paths(el):
    result=[]; box=el.get('clip'); filled=el.get('fill','none') != 'none'
    for path in el['paths']:
        pts=path['points']; closed=path.get('closed',False)
        if filled:
            if not closed or not convex(pts): raise ValueError('Filled paths must be closed convex polygons')
            clipped=polygon(pts,box)
            if len(clipped)>=3: result.append({'points':clipped,'closed':True,'filled':True,'stroke':box is None})
            # Clip the original outline separately: polygon clipping must not invent
            # visible strokes along the panel's newly introduced cut edges.
            if box is not None and el.get('stroke','#000000')!='none':
                for a,b in zip(pts,pts[1:]+pts[:1]):
                    part=segment(a,b,box)
                    if part is not None:result.append({'points':part,'closed':False,'filled':False,'stroke':True})
        else:
            chain=[]
            seq=pts+[pts[0]] if closed else pts
            for a,b in zip(seq,seq[1:]):
                part=segment(a,b,box)
                if part is None:
                    if chain: result.append({'points':chain,'closed':False}); chain=[]
                elif chain and all(abs(chain[-1][j]-part[0][j])<1e-9 for j in (0,1)):
                    chain.append(part[1])
                else:
                    if chain: result.append({'points':chain,'closed':False})
                    chain=part
            if chain: result.append({'points':chain,'closed':False})
    return result

def validate_scene(scene):
    fields(scene,'width height elements')
    for k in ('width','height'):
        if number(scene[k]) <= 0 or scene[k] > 4032: raise ValueError('Slide dimensions must be in (0, 4032] points')
    for el in scene['elements']:
        kind=el['kind']
        common='kind name '
        if kind=='shape':
            fields(el,common+'paths fill stroke lineWidth dash clip')
            color(el.get('fill','none')); color(el.get('stroke','#000000'))
            if number(el.get('lineWidth',1))<0: raise ValueError('Negative line width')
            if el.get('clip') is not None:
                b=el['clip']
                if len(b)!=4 or any(number(v) is None for v in b) or b[0]>=b[2] or b[1]>=b[3]: raise ValueError('Invalid clip rectangle')
            for path in el['paths']:
                fields(path,'points closed')
                if len(path['points'])<2: raise ValueError('Path must contain at least two points')
                for pt in path['points']:
                    if len(pt)!=2: raise ValueError('Expected x,y pairs')
                    for v in pt: number(v)
        elif kind in ('text','image'):
            fields(el,common+'x y width height '+('text runs size font color alignment rotation' if kind=='text' else 'file'))
            for k in ('x','y','width','height'): number(el[k])
            if el['width']<=0 or el['height']<=0: raise ValueError('Non-positive object dimensions')
            if kind=='text':
                if el.get('alignment','left') not in ('left','center','right'): raise ValueError('Unsupported alignment')
                number(el.get('rotation',0)); color(el.get('color','#000000'))
                if not 1 <= number(el.get('size',12)) <= 4000: raise ValueError('Invalid text size')
                if ('text' in el)==('runs' in el): raise ValueError('Supply exactly one of text or runs')
                for run in el.get('runs',[]):
                    fields(run,'text style'); st=run.get('style',{})
                    fields(st,'size font color bold italic baseline')
                    if not 1 <= number(st.get('size',el.get('size',12))) <= 4000: raise ValueError('Invalid run size')
                    color(st.get('color',el.get('color','#000000')))
                    number(st.get('baseline',0))
        else: raise ValueError('Unsupported kind: '+str(kind))

def fill(parent, value):
    if value=='none': node('a:noFill',parent=parent)
    else: node('a:srgbClr',{'val':value[1:]},node('a:solidFill',parent=parent))

def transform(parent,x,y,w,h,rotation=0):
    tr=node('a:xfrm',{'rot':round(rotation*60000)} if rotation else {},parent)
    node('a:off',{'x':emu(x),'y':emu(y)},tr)
    node('a:ext',{'cx':max(1,emu(w)),'cy':max(1,emu(h))},tr)

def nonvisual(parent,i,name,picture=False,text=False):
    nv=node('p:nvPicPr' if picture else 'p:nvSpPr',parent=parent)
    node('p:cNvPr',{'id':i,'name':name},nv)
    node('p:cNvPicPr' if picture else 'p:cNvSpPr',{'txBox':'1'} if text else {},nv)
    node('p:nvPr',parent=nv)

def group(parent):
    tree=node('p:spTree',parent=parent)
    nv=node('p:nvGrpSpPr',parent=tree);node('p:cNvPr',{'id':1,'name':''},nv)
    node('p:cNvGrpSpPr',parent=nv);node('p:nvPr',parent=nv)
    gp=node('p:grpSpPr',parent=tree);tr=node('a:xfrm',parent=gp)
    for tag,attr in [('off',{'x':0,'y':0}),('ext',{'cx':0,'cy':0}),('chOff',{'x':0,'y':0}),('chExt',{'cx':0,'cy':0})]:node('a:'+tag,attr,tr)
    return tree

def relationships(entries):
    root=ET.Element('Relationships',xmlns=REL)
    for ident,kind,target in entries:
        ET.SubElement(root,'Relationship',Id=ident,Type=NS['r']+'/'+kind,Target=target)
    return xml(root)

def build(scene,base,output):
    validate_scene(scene); output=Path(output)
    if output.exists(): raise FileExistsError('Refusing to overwrite '+str(output))
    slide=node('p:sld'); tree=group(node('p:cSld',parent=slide)); media={}
    rels=[('rId1','slideLayout','../slideLayouts/slideLayout1.xml')]
    counts={'textboxes':0,'vector_shapes':0,'images':0}
    for i,original in enumerate(scene['elements'],2):
        el=copy.deepcopy(original); kind=el['kind']; name=el.get('name',kind)+' ['+str(i)+']'
        if kind=='shape':
            paths=clipped_paths(el)
            if not paths: continue
            pts=[pt for path in paths for pt in path['points']]
            x=min(p[0] for p in pts);y=min(p[1] for p in pts)
            w=max(.001,max(p[0] for p in pts)-x);h=max(.001,max(p[1] for p in pts)-y)
            sp=node('p:sp',parent=tree);nonvisual(sp,i,name)
            pr=node('p:spPr',parent=sp);transform(pr,x,y,w,h)
            geo=node('a:custGeom',parent=pr)
            for tag in ('avLst','gdLst','ahLst','cxnLst'):node('a:'+tag,parent=geo)
            node('a:rect',{'l':'l','t':'t','r':'r','b':'b'},geo)
            pl=node('a:pathLst',parent=geo)
            for path in paths:
                pp=node('a:path',{'w':max(1,emu(w)),'h':max(1,emu(h)),'fill':'norm' if path.get('filled',False) else 'none','stroke':int(path.get('stroke',True))},pl)
                for j,(xx,yy) in enumerate(path['points']):
                    node('a:pt',{'x':emu(xx-x),'y':emu(yy-y)},node('a:lnTo' if j else 'a:moveTo',parent=pp))
                if path['closed']:node('a:close',parent=pp)
            fill(pr,el.get('fill','none'))
            ln=node('a:ln',{'w':emu(el.get('lineWidth',1))},pr);fill(ln,el.get('stroke','#000000'))
            if el.get('dash'):node('a:prstDash',{'val':'dash'},ln)
            counts['vector_shapes']+=1
        elif kind=='text':
            sp=node('p:sp',parent=tree);nonvisual(sp,i,name,text=True)
            pr=node('p:spPr',parent=sp);transform(pr,el['x'],el['y'],el['width'],el['height'],el.get('rotation',0))
            node('a:avLst',parent=node('a:prstGeom',{'prst':'rect'},pr));fill(pr,'none');fill(node('a:ln',parent=pr),'none')
            tx=node('p:txBody',parent=sp)
            bp=node('a:bodyPr',{'wrap':'none','lIns':0,'rIns':0,'tIns':0,'bIns':0,'anchor':'t'},tx);node('a:noAutofit',parent=bp)
            node('a:lstStyle',parent=tx);para=node('a:p',parent=tx)
            node('a:pPr',{'algn':{'left':'l','center':'ctr','right':'r'}[el.get('alignment','left')]},para)
            runs=el.get('runs',[{'text':el.get('text',''),'style':{}}])
            for run in runs:
                st=run.get('style',{});r=node('a:r',parent=para)
                rp=node('a:rPr',{'lang':'en-US','sz':round(st.get('size',el.get('size',12))*100),'b':int(st.get('bold',False)),'i':int(st.get('italic',False)),'baseline':round(st.get('baseline',0))},r)
                fill(rp,st.get('color',el.get('color','#000000')))
                for tag in ('latin','ea','cs'):node('a:'+tag,{'typeface':st.get('font',el.get('font','Arial'))},rp)
                node('a:t',parent=r,text=run['text'])
            node('a:endParaRPr',{'lang':'en-US','sz':round(el.get('size',12)*100)},para);counts['textboxes']+=1
        else:
            file=Path(el['file'])
            if not file.is_absolute():file=Path(base)/file
            data=file.read_bytes();ext=file.suffix.lower().lstrip('.')
            if ext not in ('png','jpg','jpeg'):raise ValueError('Only PNG/JPEG images are supported')
            if not (data.startswith(b'\x89PNG\r\n\x1a\n') if ext=='png' else data.startswith(b'\xff\xd8\xff')):raise ValueError('Image signature does not match extension')
            dest='image'+str(i)+'.'+ext;media['ppt/media/'+dest]=data;rid='rId'+str(len(rels)+1)
            rels.append((rid,'image','../media/'+dest))
            pic=node('p:pic',parent=tree);nonvisual(pic,i,name,picture=True)
            bf=node('p:blipFill',parent=pic);node('a:blip',{'{'+NS['r']+'}embed':rid},bf)
            node('a:fillRect',parent=node('a:stretch',parent=bf))
            pr=node('p:spPr',parent=pic);transform(pr,el['x'],el['y'],el['width'],el['height'])
            node('a:avLst',parent=node('a:prstGeom',{'prst':'rect'},pr));counts['images']+=1
    node('a:masterClrMapping',parent=node('p:clrMapOvr',parent=slide))
    master=node('p:sldMaster');group(node('p:cSld',parent=master))
    node('p:clrMap',dict(bg1='lt1',tx1='dk1',bg2='lt2',tx2='dk2',accent1='accent1',accent2='accent2',accent3='accent3',accent4='accent4',accent5='accent5',accent6='accent6',hlink='hlink',folHlink='folHlink'),master)
    node('p:sldLayoutId',{'id':2147483649,'{'+NS['r']+'}id':'rId1'},node('p:sldLayoutIdLst',parent=master))
    txs=node('p:txStyles',parent=master)
    for tag in ('titleStyle','bodyStyle','otherStyle'):node('p:'+tag,parent=txs)
    layout=node('p:sldLayout',{'type':'blank','preserve':'1'});group(node('p:cSld',{'name':'Blank'},layout));node('a:masterClrMapping',parent=node('p:clrMapOvr',parent=layout))
    presentation=node('p:presentation');node('p:sldMasterId',{'id':2147483648,'{'+NS['r']+'}id':'rId1'},node('p:sldMasterIdLst',parent=presentation))
    node('p:sldId',{'id':256,'{'+NS['r']+'}id':'rId2'},node('p:sldIdLst',parent=presentation))
    node('p:sldSz',{'cx':emu(scene['width']),'cy':emu(scene['height']),'type':'custom'},presentation)
    node('p:notesSz',{'cx':6858000,'cy':9144000},presentation)
    theme=node('a:theme',{'name':'Scientific figure'});te=node('a:themeElements',parent=theme)
    cs=node('a:clrScheme',{'name':'Scientific'},te)
    for tag,c in dict(dk1='000000',lt1='FFFFFF',dk2='222222',lt2='EEEEEE',accent1='0072B2',accent2='D55E00',accent3='009E73',accent4='CC79A7',accent5='E69F00',accent6='56B4E9',hlink='0000FF',folHlink='800080').items():node('a:srgbClr',{'val':c},node('a:'+tag,parent=cs))
    fs=node('a:fontScheme',{'name':'Arial'},te)
    for tag in ('majorFont','minorFont'):
        f=node('a:'+tag,parent=fs)
        for sub in ('latin','ea','cs'):node('a:'+sub,{'typeface':'Arial'},f)
    fmt=node('a:fmtScheme',{'name':'Scientific'},te)
    for tag in ('fillStyleLst','bgFillStyleLst'):
        lst=node('a:'+tag,parent=fmt)
        for _ in range(3):node('a:schemeClr',{'val':'phClr'},node('a:solidFill',parent=lst))
    lines=node('a:lnStyleLst',parent=fmt)
    for w in (6350,12700,19050):
        ln=node('a:ln',{'w':w},lines);node('a:schemeClr',{'val':'phClr'},node('a:solidFill',parent=ln));node('a:prstDash',{'val':'solid'},ln)
    effects=node('a:effectStyleLst',parent=fmt)
    for _ in range(3):node('a:effectLst',parent=node('a:effectStyle',parent=effects))
    # Required schema order: fills, lines, effects, background fills.
    bg=fmt.find('a:bgFillStyleLst',NS);fmt.remove(bg);fmt.append(bg)
    package={'ppt/presentation.xml':xml(presentation),'ppt/slides/slide1.xml':xml(slide),'ppt/slideMasters/slideMaster1.xml':xml(master),'ppt/slideLayouts/slideLayout1.xml':xml(layout),'ppt/theme/theme1.xml':xml(theme),
             '_rels/.rels':relationships([('rId1','officeDocument','ppt/presentation.xml')]),
             'ppt/_rels/presentation.xml.rels':relationships([('rId1','slideMaster','slideMasters/slideMaster1.xml'),('rId2','slide','slides/slide1.xml')]),
             'ppt/slides/_rels/slide1.xml.rels':relationships(rels),
             'ppt/slideLayouts/_rels/slideLayout1.xml.rels':relationships([('rId1','slideMaster','../slideMasters/slideMaster1.xml')]),
             'ppt/slideMasters/_rels/slideMaster1.xml.rels':relationships([('rId1','slideLayout','../slideLayouts/slideLayout1.xml'),('rId2','theme','../theme/theme1.xml')]),**media}
    types=ET.Element('Types',xmlns=CT)
    for ext,mime in [('rels','application/vnd.openxmlformats-package.relationships+xml'),('xml','application/xml'),('png','image/png'),('jpg','image/jpeg'),('jpeg','image/jpeg')]:ET.SubElement(types,'Default',Extension=ext,ContentType=mime)
    for path,typ in [('ppt/presentation.xml','presentation.main'),('ppt/slides/slide1.xml','slide'),('ppt/slideMasters/slideMaster1.xml','slideMaster'),('ppt/slideLayouts/slideLayout1.xml','slideLayout')]:ET.SubElement(types,'Override',PartName='/'+path,ContentType='application/vnd.openxmlformats-officedocument.presentationml.'+typ+'+xml')
    ET.SubElement(types,'Override',PartName='/ppt/theme/theme1.xml',ContentType='application/vnd.openxmlformats-officedocument.theme+xml')
    package['[Content_Types].xml']=xml(types)
    output.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(output,'x',ZIP_DEFLATED) as z:
        for path,data in package.items():z.writestr(path,data)
    return counts

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('scene',type=Path);parser.add_argument('output',type=Path);args=parser.parse_args()
    result=build(json.loads(args.scene.read_text(encoding='utf-8')),args.scene.resolve().parent,args.output)
    print(json.dumps(result,indent=2))
