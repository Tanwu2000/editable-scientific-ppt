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
