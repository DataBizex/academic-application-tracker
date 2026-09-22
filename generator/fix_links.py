"""Make every in-workbook hyperlink a pure internal link (location only, no external file target)."""
import re, sys, zipfile, shutil, os, tempfile

def fix(path):
    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(path) as z:
        z.extractall(tmp)
        names = z.namelist()
    ws_dir = os.path.join(tmp, "xl", "worksheets")
    fixed = 0
    for fn in os.listdir(ws_dir):
        if not fn.endswith(".xml"):
            continue
        sp = os.path.join(ws_dir, fn)
        rp = os.path.join(ws_dir, "_rels", fn + ".rels")
        xml = open(sp, encoding="utf-8").read()
        drop_ids = []
        def repl(m):
            nonlocal fixed
            tag = m.group(0)
            if 'location="' not in tag:
                return tag  # genuine external link (website, LinkedIn)
            rid = re.search(r'r:id="([^"]+)"', tag)
            if rid:
                drop_ids.append(rid.group(1))
            tag = re.sub(r'\s+r:id="[^"]*"', "", tag)
            tag = re.sub(r'\s+display="[^"]*"', "", tag)
            fixed += 1
            return tag
        xml = re.sub(r'<hyperlink [^>]*/>', repl, xml)
        open(sp, "w", encoding="utf-8").write(xml)
        if drop_ids and os.path.exists(rp):
            rels = open(rp, encoding="utf-8").read()
            for rid in drop_ids:
                rels = re.sub(r'<Relationship Id="%s"[^>]*/>' % re.escape(rid), "", rels)
            open(rp, "w", encoding="utf-8").write(rels)
    out = path + ".tmp"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(os.path.join(tmp, n), n)
    shutil.move(out, path)
    shutil.rmtree(tmp)
    print(f"{os.path.basename(path)}: {fixed} internal links fixed")

for p in sys.argv[1:]:
    fix(p)
