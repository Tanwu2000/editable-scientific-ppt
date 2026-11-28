import copy
import json
from pathlib import Path
import sys
import tempfile
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'));sys.path.insert(0,str(ROOT/'tools'))
from native_pptx import build, clipped_paths, segment, polygon, validate_scene, NS
from inspect_pptx import inspect
from install_local import install
from figure_style import prepare_scene

class NativeTests(unittest.TestCase):
    def setUp(self):self.scene=json.loads((ROOT/'examples/minimal_scene.json').read_text(encoding='utf-8'))

    def test_line_clipping_crossing_and_outside(self):
        self.assertEqual(segment([-1,5],[11,5],[0,0,10,10]),[[0.,5.],[10.,5.]])
        self.assertIsNone(segment([-1,-2],[11,-2],[0,0,10,10]))
        self.assertEqual(segment([5,-1],[5,11],[0,0,10,10]),[[5.,0.],[5.,10.]])

    def test_clipped_polygon_stays_inside(self):
        pts=polygon([[-2,3],[3,-2],[8,3],[3,8]],[0,0,6,6])
        self.assertTrue(all(0<=x<=6 and 0<=y<=6 for x,y in pts))
        self.assertGreaterEqual(len(pts),4)

    def test_outside_excursion_does_not_bridge(self):
        shape={'paths':[{'points':[[1,1],[-2,1],[-2,9],[1,9]]}],'clip':[0,0,10,10]}
        paths=clipped_paths(shape)
        self.assertEqual(len(paths),2)
        self.assertEqual(paths[0]['points'][-1],[0.,1.])
        self.assertEqual(paths[1]['points'][0],[0.,9.])

    def test_native_package_unicode_baseline_and_clipping(self):
        with tempfile.TemporaryDirectory() as tmp:
            out=Path(tmp)/'figure.pptx';build(self.scene,ROOT/'examples',out);report=inspect(out)
            self.assertEqual((report['textboxes'],report['vector_shapes'],report['images']),(5,3,0))
            self.assertIn('Displacement, δ (µm)',report['texts'])
            self.assertEqual(report['fonts'],['Arial'])
            with ZipFile(out) as z:
                slide=ET.fromstring(z.read('ppt/slides/slide1.xml'))
                self.assertTrue(any(e.get('baseline')=='-25000' for e in slide.findall('.//a:rPr',NS)))
                for el in self.scene['elements']:
                    if el['kind']=='shape' and el.get('clip'):
                        b=el['clip']
                        for path in clipped_paths(el):
                            self.assertTrue(all(b[0]<=x<=b[2] and b[1]<=y<=b[3] for x,y in path['points']))
            with self.assertRaises(FileExistsError):build(self.scene,ROOT/'examples',out)

    def test_unsupported_properties_fail(self):
        self.scene['elements'][0]['opacity']=0.2
        with self.assertRaises(ValueError):validate_scene(self.scene)

    def test_nested_group_identity_bounds_and_unique_ids(self):
        # Non-origin panel coordinates and rotated labels must not move on grouping.
        self.scene['elements']=[{'kind':'group','name':'Panel','elements':self.scene['elements']}]
        with tempfile.TemporaryDirectory() as tmp:
            out=Path(tmp)/'group.pptx';build(self.scene,ROOT/'examples',out)
            report=inspect(out)
            self.assertEqual((report['groups'],report['textboxes'],report['vector_shapes']),(1,5,3))
            self.assertEqual((report['grouping_locks'],report['placeholders']),(0,0))
            with ZipFile(out) as z:
                slide=ET.fromstring(z.read('ppt/slides/slide1.xml'))
                for tr in slide.findall('.//p:grpSp/p:grpSpPr/a:xfrm',NS):
                    self.assertEqual(tr.find('a:off',NS).attrib,tr.find('a:chOff',NS).attrib)
                    self.assertEqual(tr.find('a:ext',NS).attrib,tr.find('a:chExt',NS).attrib)
                    self.assertGreater(int(tr.find('a:ext',NS).get('cx')),0)
                # Child coordinates survive exactly; group bounds include rotated text.
                label=next(e for e in slide.findall('.//p:sp',NS) if 'y-axis' in e.find('p:nvSpPr/p:cNvPr',NS).get('name'))
                self.assertEqual(int(label.find('p:spPr/a:xfrm/a:off',NS).get('x')),-35*12700)

    def test_title_roles_helper_optin_and_explicit_style(self):
        original=copy.deepcopy(self.scene)
        title=self.scene['elements'][0]
        title.update(role='title',centerOn=[60,320],x=2,width=10,alignment='left')
        self.scene['elements'] += [{'kind':'text','name':'helper','role':'helper','x':0,'y':0,'width':100,'height':20,'text':'Generated with exporter'},
                                   {'kind':'text','name':'science-note','x':0,'y':0,'width':100,'height':20,'text':'n = 12'}]
        ready=prepare_scene(self.scene)
        self.assertEqual((ready['elements'][0]['x'],ready['elements'][0]['width'],ready['elements'][0]['alignment']),(60,260,'center'))
        self.assertNotIn('helper',[e['name'] for e in ready['elements']])
        self.assertIn('science-note',[e['name'] for e in ready['elements']])
        self.assertEqual(self.scene['elements'][0]['x'],2) # input not mutated
        self.assertEqual(ready['elements'][1]['stroke'],original['elements'][1]['stroke'])
        self.scene['showHelperText']=True
        self.assertIn('helper',[e['name'] for e in prepare_scene(self.scene)['elements']])

    def test_line_role_defaults_preserve_explicit_overrides(self):
        template={'kind':'shape','paths':[{'points':[[0,0],[10,10]]}]}
        scene={'width':100,'height':100,'elements':[dict(template,role=r) for r in ['axis','border','curve','mesh']]}
        ready=prepare_scene(scene)
        self.assertEqual([e['lineWidth'] for e in ready['elements']],[1.2,1.2,1.5,.7])
        self.assertTrue(all(e['stroke']=='#000000' for e in ready['elements']))
        scene['elements'][0].update(stroke='#123456',lineWidth=2)
        self.assertEqual(prepare_scene(scene)['elements'][0]['stroke'],'#123456')
        self.assertEqual(prepare_scene(scene)['elements'][0]['lineWidth'],2)

    def test_flat_optout_and_fully_clipped_group(self):
        self.scene['group']=False
        with tempfile.TemporaryDirectory() as tmp:
            out=Path(tmp)/'flat.pptx';build(self.scene,ROOT/'examples',out)
            self.assertEqual(inspect(out)['groups'],0)
            self.scene['group']=True
            self.scene['elements']=[{'kind':'group','elements':[{'kind':'shape','clip':[0,0,5,5],'paths':[{'points':[[10,10],[20,20]]}]}]}]
            out=Path(tmp)/'empty.pptx';build(self.scene,ROOT/'examples',out)
            self.assertEqual(inspect(out)['groups'],0)

    def test_filled_clip_does_not_stroke_artificial_boundary(self):
        paths=clipped_paths({'fill':'#EEEEEE','stroke':'#000000','clip':[0,0,5,5],
            'paths':[{'points':[[-1,-1],[6,-1],[6,6],[-1,6]],'closed':True}]})
        self.assertEqual(len(paths),1)
        self.assertTrue(paths[0]['filled'])
        self.assertFalse(paths[0]['stroke'])

    def test_embedded_raster_has_internal_relationship(self):
        import struct, zlib
        def chunk(tag,data):return struct.pack('!I',len(data))+tag+data+struct.pack('!I',zlib.crc32(tag+data)&0xffffffff)
        png=b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('!2I5B',1,1,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(b'\x00\x00\x72\xb2'))+chunk(b'IEND',b'')
        with tempfile.TemporaryDirectory() as tmp:
            base=Path(tmp);(base/'pixel.png').write_bytes(png)
            self.scene['elements'].insert(0,{'kind':'image','name':'synthetic-field','file':'pixel.png','x':60,'y':60,'width':260,'height':145})
            out=base/'figure.pptx';build(self.scene,base,out)
            self.assertEqual(inspect(out)['images'],1)
            with ZipFile(out) as z:self.assertTrue(any(p.startswith('ppt/media/') for p in z.namelist()))

    def test_install_update_backup_and_git_guard(self):
        with tempfile.TemporaryDirectory() as tmp:
            base=Path(tmp);source=base/'source';source.mkdir();(source/'SKILL.md').write_text('version one')
            dest=base/'skills';res=install(source,dest);target=Path(res['installed'])
            with self.assertRaises(FileExistsError):install(source,dest)
            (source/'SKILL.md').write_text('version two');res=install(source,dest,True)
            self.assertEqual((Path(res['backup'])/'SKILL.md').read_text(),'version one')
            self.assertEqual((target/'SKILL.md').read_text(),'version two')
            (target/'.git').mkdir()
            with self.assertRaises(ValueError):install(source,dest,True)

if __name__=='__main__':unittest.main()
