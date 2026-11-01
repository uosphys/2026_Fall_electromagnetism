# LibreOffice로 PPTX → PDF 내보내기 (한글·라틴 사이 자동 간격 끄기)
import sys, time, subprocess, os, tempfile, uno
from com.sun.star.beans import PropertyValue
sys.path.insert(0, os.path.join(sys.argv[3], "scripts"))
from office.soffice import get_soffice_env
src, dst = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
prof = tempfile.mkdtemp(prefix="lo_")
p = subprocess.Popen(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--invisible", "--norestore",
    "--accept=pipe,name=emw6;urp;"], env=get_soffice_env())
ctx = None
for _ in range(60):
    try:
        local = uno.getComponentContext()
        res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = res.resolve("uno:pipe,name=emw6;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
smgr = ctx.ServiceManager
desk = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
def pv(n, v): x = PropertyValue(); x.Name = n; x.Value = v; return x
doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))
n = 0
def walk(shape):
    global n
    if shape.supportsService("com.sun.star.drawing.GroupShape"):
        for i in range(shape.getCount()): walk(shape.getByIndex(i))
        return
    if hasattr(shape, "getText"):
        try:
            en = shape.getText().createEnumeration()
            while en.hasMoreElements():
                par = en.nextElement()
                par.setPropertyValue("ParaIsCharacterDistance", False); n += 1
        except Exception: pass
pages = doc.getDrawPages()
for i in range(pages.getCount()):
    pg = pages.getByIndex(i)
    for j in range(pg.getCount()): walk(pg.getByIndex(j))
doc.storeToURL(uno.systemPathToFileUrl(dst), (pv("FilterName", "impress_pdf_Export"),))
doc.close(True)
try: desk.terminate()
except Exception: pass
p.wait(timeout=30)
print("paragraphs", n, "→", dst)
