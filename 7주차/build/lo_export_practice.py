"""LibreOffice UNO로 PPTX → PDF. 한글-영문 사이 자동 간격(ParaIsCharacterDistance)을 끈 뒤 내보낸다."""
import subprocess, sys, time, uno
from com.sun.star.beans import PropertyValue
src, dst = sys.argv[1], sys.argv[2]
p = subprocess.Popen(["soffice", "-env:UserInstallation=file:///tmp/em-lo-uno", "--headless", "--invisible",
                      "--accept=socket,host=localhost,port=2002;urp;"])
ctx = None
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(60):
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(0.5)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
def pv(n, v):
    x = PropertyValue(); x.Name = n; x.Value = v; return x
doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))
n = 0
pages = doc.DrawPages
for i in range(pages.Count):
    pg = pages.getByIndex(i)
    for j in range(pg.Count):
        sh = pg.getByIndex(j)
        if sh.supportsService("com.sun.star.drawing.Text"):
            en = sh.getText().createEnumeration()
            while en.hasMoreElements():
                para = en.nextElement()
                try:
                    para.ParaIsCharacterDistance = False; n += 1
                except Exception: pass
doc.storeToURL(uno.systemPathToFileUrl(dst), (pv("FilterName", "impress_pdf_Export"),))
doc.close(True)
try: desktop.terminate()
except Exception: pass
p.wait(timeout=30)
print("paragraphs patched:", n)
