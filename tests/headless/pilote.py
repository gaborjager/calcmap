"""Pilote LibreOffice sans écran : ouvre un classeur, lance une macro de test
de l'extension, exporte en PNG la 1re forme de la 1re feuille (s'il y en a).

Usage : python3 pilote.py <profil> <classeur.ods> <macro> <image.png> [port]
(appelé par lancer.sh)
"""
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(nom, valeur):
    p = PropertyValue()
    p.Name = nom
    p.Value = valeur
    return p


profil, classeur, macro, image = sys.argv[1:5]
port = sys.argv[5] if len(sys.argv) > 5 else "2120"
proc = subprocess.Popen(["soffice", f"-env:UserInstallation=file://{profil}", "--headless", "--norestore",
                         f"--accept=socket,host=localhost,port={port};urp;"])
desk = None
try:
    ctx = uno.getComponentContext()
    res = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
    for _ in range(60):
        try:
            rctx = res.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(1)
    smgr = rctx.ServiceManager
    desk = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", rctx)
    # MacroExecutionMode 4 : comme une ouverture normale (événement « Ouvrir le document »)
    doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(classeur), "_blank", 0,
                                    (prop("MacroExecutionMode", 4),))
    sp = smgr.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory",
                                        rctx).createScriptProvider("")
    sp.getScript(f"vnd.sun.star.script:CalcMap.CalcMap.{macro}?language=Basic&location=application"
                 ).invoke((), (), ())
    page = doc.Sheets.getByIndex(0).DrawPage
    if page.Count:
        f = smgr.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", rctx)
        f.setSourceDocument(page.getByIndex(0))
        fd = uno.Any("[]com.sun.star.beans.PropertyValue", (prop("PixelWidth", 1400),))
        f.filter((prop("URL", uno.systemPathToFileUrl(image)), prop("MediaType", "image/png"),
                  prop("FilterData", fd)))
    doc.close(True)
finally:
    try:
        desk.terminate()
    except Exception:
        pass
    time.sleep(2)
    proc.kill()
