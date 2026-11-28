"""Install this local skill copy. Existing copies require --update and are backed up."""
import argparse
from datetime import datetime, timezone
import os
from pathlib import Path
import shutil
import tempfile

NAME='editable-scientific-ppt'
ALLOW=('SKILL.md','VERSION','README.md','CHANGELOG.md','LICENSE','agents','references','scripts','examples','tools')

def install(source,dest,update=False):
    source=Path(source).resolve();dest=Path(dest).expanduser().resolve();target=dest/NAME
    if not (source/'SKILL.md').is_file():raise ValueError('Missing source SKILL.md')
    if target.is_symlink():raise ValueError('Refusing to replace a linked skill; update its source instead')
    if target.exists():
        if (target/'.git').exists():raise ValueError('Git installation: review changes and use git pull --ff-only')
        if not update:raise FileExistsError('Already installed; use --update after reviewing the new version')
        if not target.is_dir():raise ValueError('Target is not a directory')
        if source==target or source.is_relative_to(target):raise ValueError('Run update from a separate source checkout')
    # Detect junctions and other redirection without recursively moving an unchecked target.
    if target.exists() and target.resolve()!=target:raise ValueError('Refusing redirected installation path')
    dest.mkdir(parents=True,exist_ok=True)
    stage=Path(tempfile.mkdtemp(prefix=NAME+'-stage-',dir=dest.parent));backup=None
    try:
        for name in ALLOW:
            p=source/name
            if not p.exists():continue
            if p.is_dir():shutil.copytree(p,stage/name,ignore=shutil.ignore_patterns('__pycache__','*.pyc'))
            else:shutil.copy2(p,stage/name)
        if target.exists():
            backups=dest.parent/'skill-backups';backups.mkdir(exist_ok=True)
            backup=backups/(NAME+'-'+datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ'))
            if target.resolve().parent!=dest:raise ValueError('Target escaped intended skill directory')
            target.rename(backup)
        try:stage.rename(target)
        except Exception:
            if backup is not None:backup.rename(target)
            raise
        return {'installed':str(target),'backup':str(backup) if backup else None}
    finally:
        # Only our freshly created stage is removed; never recursively delete an installed skill.
        if stage.exists() and stage.parent==dest.parent and stage.name.startswith(NAME+'-stage-'):
            shutil.rmtree(stage)

if __name__=='__main__':
    import sys
    if hasattr(sys.stdout,'reconfigure'): sys.stdout.reconfigure(encoding='utf-8')
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--dest',type=Path,default=Path(os.environ.get('CODEX_HOME',str(Path.home()/'.codex')))/'skills')
    ap.add_argument('--update',action='store_true');a=ap.parse_args()
    import json
    print(json.dumps(install(Path(__file__).resolve().parents[1],a.dest,a.update),indent=2,ensure_ascii=False))
