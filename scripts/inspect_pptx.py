"""Inspect native editability and package links; this is not a visual QA test."""
import argparse
import json
import posixpath
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from native_pptx import NS

def inspect(path):
    with ZipFile(path) as z:
        if z.testzip(): raise ValueError('Corrupt archive')
        names=set(z.namelist()); roots={}
        for name in names:
            if name.endswith(('.xml','.rels')): roots[name]=ET.fromstring(z.read(name))
        for name,root in roots.items():
            if not name.endswith('.rels'):continue
            base='' if name=='_rels/.rels' else posixpath.dirname(posixpath.dirname(name))
            for rel in root:
                if rel.get('TargetMode')=='External':raise ValueError('External link: '+rel.get('Target',''))
                target=posixpath.normpath(posixpath.join(base,rel.get('Target','')))
                if target not in names:raise ValueError('Missing relationship target: '+target)
        slides=[root for name,root in roots.items() if name.startswith('ppt/slides/slide') and name.endswith('.xml')]
        size=roots['ppt/presentation.xml'].find('p:sldSz',NS)
        result={'slides':len(slides),'width_pt':int(size.get('cx'))/12700,'height_pt':int(size.get('cy'))/12700,'textboxes':0,'vector_shapes':0,'images':0,'fonts':set(),'texts':[],'external_links':0,'groups':0,'grouping_locks':0,'placeholders':0}
        for slide in slides:
            result['groups']+=len(slide.findall('.//p:grpSp',NS))
            result['placeholders']+=len(slide.findall('.//p:ph',NS))
            result['grouping_locks']+=sum(e.get('noGrp') in ('1','true') for e in slide.iter())
            ids=[e.get('id') for e in slide.findall('.//p:cNvPr',NS)]
            if len(ids)!=len(set(ids)):raise ValueError('Duplicate shape IDs')
            result['images']+=len(slide.findall('.//p:pic',NS))
            for sp in slide.findall('.//p:sp',NS):
                text=sp.find('p:txBody',NS)
                if text is not None:
                    result['textboxes']+=1
                    result['texts'].append(''.join(e.text or '' for e in text.findall('.//a:t',NS)))
                    for f in text.findall('.//a:latin',NS):result['fonts'].add(f.get('typeface'))
                if sp.find('p:spPr/a:custGeom',NS) is not None:result['vector_shapes']+=1
        result['fonts']=sorted(result['fonts'])
        return result

if __name__=='__main__':
    import sys
    if hasattr(sys.stdout,'reconfigure'): sys.stdout.reconfigure(encoding='utf-8')
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('pptx',type=Path);a=ap.parse_args()
    print(json.dumps(inspect(a.pptx),indent=2,ensure_ascii=False))
