"""Apply explicit scientific-figure roles; never infer or delete scientific text."""
import copy

ROLES = {'title', 'axis', 'border', 'curve', 'mesh', 'helper', 'label'}

def prepare_scene(source):
    scene = copy.deepcopy(source)
    def prepare(elements):
        out = []
        for el in elements:
            role = el.get('role')
            if role == 'helper' and not scene.get('showHelperText', False):
                continue
            if el['kind'] == 'group':
                el['elements'] = prepare(el['elements'])
                if not el['elements']: continue
            elif el['kind'] == 'text':
                el.setdefault('font', 'Arial')
                el.setdefault('color', '#000000')
                if role == 'title':
                    left, right = el['centerOn']
                    el.update(x=left, width=right-left, alignment='center')
            elif el['kind'] == 'shape':
                el.setdefault('stroke', '#000000')
                el.setdefault('lineWidth', {'axis':1.2, 'border':1.2,
                                           'curve':1.5, 'mesh':0.7}.get(role, 1.0))
            out.append(el)
        return out
    scene['elements'] = prepare(scene['elements'])
    return scene
