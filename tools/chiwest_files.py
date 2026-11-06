#!/usr/bin/env python3
"""Keep Chi West's copies of the case sheet images in step with the Graft Tracker app (AJ 11/5: every case and PO in
Chi West opens its case sheet).

The Graft Tracker app (GT) holds each case's files (case sheet, transport log) as its own assets, listed by file name in
its db doc meta/assets {files: {name: asset id}}; GT's months docs list them per case (cases[].files[].f). Chi West can
only show images from its own asset store, so the .jpg files are copied there and Chi West's db doc data/files maps
{files: {name: {id: Chi West asset id, src: GT asset id}}}. A file re-made in GT (e.g. a PO added to its banner) gets
a new GT asset id, so it is copied again.

    python3 -I tools/chiwest_files.py todo GTDB/meta/assets.json CW/data/files.json
        prints the GT asset ids still to copy, in groups of 10 (missing CW/data/files.json = copy everything)
    python3 -I tools/chiwest_files.py record CW/data/files.json GTDB/meta/assets.json OUT.json NEW_ID=SRC_ID ...
        adds the copies to the map and writes the new data/files doc to OUT.json

Copy each group with the Artifact tool: publish, url = Chi West, asset true, from_url = GT, asset_ids = the group. Its
result lists "/_blob/<new id> ... source id <src id>" per file; pass those pairs to `record`, then ArtifactData set Chi
West data/files from OUT.json (pin if_version once the doc exists).
"""
import datetime as dt
import json
import os
import sys


def load(p):
    if not p or not os.path.exists(p):
        return {}
    d = json.load(open(p))
    return d.get('data', d) if isinstance(d.get('data'), dict) else d


def main():
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    cmd = sys.argv[1]
    if cmd == 'todo':
        gt = (load(sys.argv[2]).get('files') or {})
        cw = (load(sys.argv[3]).get('files') or {})
        need = [(n, a) for n, a in sorted(gt.items()) if n.lower().endswith('.jpg') and (cw.get(n) or {}).get('src') != a]
        print(f'{len(need)} file(s) to copy')
        for i in range(0, len(need), 10):
            print(json.dumps([a for _, a in need[i:i + 10]]), ' # ' + ', '.join(n for n, _ in need[i:i + 10]))
    elif cmd == 'record':
        doc = load(sys.argv[2]) or {}
        files = doc.get('files') or {}
        names = {a: n for n, a in (load(sys.argv[3]).get('files') or {}).items()}
        for arg in sys.argv[5:]:
            new, src = (x.strip().replace('/_blob/', '') for x in arg.split('=', 1))
            if src not in names:
                sys.exit(f'source id {src} is not in the GT assets list')
            files[names[src]] = {'id': new, 'src': src}
        out = {'files': files, 'about': doc.get('about') or 'Case sheet images copied from the Graft Tracker app (src) into Chi West (id), by file name',
               'updated': dt.datetime.now(dt.timezone(dt.timedelta(hours=-6))).isoformat(timespec='seconds')}
        json.dump(out, open(sys.argv[4], 'w'))
        print(f'{sys.argv[4]}: {len(files)} file(s) mapped')
    else:
        sys.exit(__doc__)


if __name__ == '__main__':
    main()
