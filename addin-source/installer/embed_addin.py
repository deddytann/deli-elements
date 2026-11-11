"""Embed a reference to the Deli Elements add-in (sideloaded via the WEF\\Developer registry key)
into a .pptx/.potx so PowerPoint loads the add-in and opens its task pane when the file opens.
Follows Microsoft's "Automatically open a task pane with a document" (storeType=Registry, store=developer).
Usage: python embed_addin.py in.potx out.potx
"""
import sys, zipfile, uuid, re

ADDIN_ID = "6f1d2b7e-3c4a-4d8e-9a5b-de11e1e2027a"
VERSION = "1.0.0.0"

WEBEXT = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{{{str(uuid.uuid5(uuid.NAMESPACE_URL, 'deli-elements-part')).upper()}}}">
  <we:reference id="{ADDIN_ID}" version="{VERSION}" store="developer" storeType="Registry"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>'''

TASKPANES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="350" row="4">
    <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
  </wetp:taskpane>
</wetp:taskpanes>'''

TP_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>'''


def embed(src, dst):
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    if "ppt/webextensions/taskpanes.xml" in names:
        raise SystemExit("already has an embedded add-in")
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for n in names:
        if n.endswith("/"):
            continue
        data = zin.read(n)
        if n == "[Content_Types].xml":
            s = data.decode("utf-8")
            add = ('<Override PartName="/ppt/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>'
                   '<Override PartName="/ppt/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>')
            s = s.replace("</Types>", add + "</Types>")
            data = s.encode("utf-8")
        elif n == "_rels/.rels":
            s = data.decode("utf-8")
            used = set(re.findall(r'Id="(rId\d+)"', s))
            rid = next(f"rId{i}" for i in range(100, 200) if f"rId{i}" not in used)
            s = s.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="ppt/webextensions/taskpanes.xml"/></Relationships>')
            data = s.encode("utf-8")
        zout.writestr(zin.getinfo(n), data)
    zout.writestr("ppt/webextensions/taskpanes.xml", TASKPANES)
    zout.writestr("ppt/webextensions/webextension1.xml", WEBEXT)
    zout.writestr("ppt/webextensions/_rels/taskpanes.xml.rels", TP_RELS)
    zout.close()


if __name__ == "__main__":
    embed(sys.argv[1], sys.argv[2])
    print("embedded add-in into", sys.argv[2])
