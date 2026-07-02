"""Strip personal metadata from the saved xlsx without disturbing Excel-authored content.

Excel stamps the signed-in user's name into the document properties on every save.
This scrubs the author fields to a generic value while leaving all worksheet content,
charts, and conditional formatting untouched (unlike an openpyxl re-save).
Run as the final step:  build -> recalc -> patch.
"""
import zipfile, re, os

HERE = os.path.dirname(os.path.abspath(__file__))
src = os.path.join(HERE, "Retirement Withdrawal Strategies Planner.xlsx")
tmp = src + ".tmp"
GENERIC = "Retirement Planner"

def clean(name, data):
    text = data.decode("utf-8")
    if name == "docProps/core.xml":
        text = re.sub(r"(<dc:creator>).*?(</dc:creator>)", r"\1" + GENERIC + r"\2", text)
        text = re.sub(r"(<cp:lastModifiedBy>).*?(</cp:lastModifiedBy>)", r"\1" + GENERIC + r"\2", text)
    if name == "docProps/app.xml":
        text = re.sub(r"(<Company>).*?(</Company>)", r"\1\2", text)
        text = re.sub(r"(<Manager>).*?(</Manager>)", r"\1\2", text)
    return text.encode("utf-8")

with zipfile.ZipFile(src, "r") as zin:
    infos = zin.infolist()
    blobs = {i.filename: zin.read(i.filename) for i in infos}

with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for i in infos:
        data = blobs[i.filename]
        if i.filename in ("docProps/core.xml", "docProps/app.xml"):
            data = clean(i.filename, data)
        zout.writestr(i, data)
os.replace(tmp, src)

# confirm the author fields now hold the generic value
with zipfile.ZipFile(src, "r") as z:
    core = z.read("docProps/core.xml").decode("utf-8")
creator = re.search(r"<dc:creator>(.*?)</dc:creator>", core)
modby = re.search(r"<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>", core)
print("metadata patched. creator:",
      creator.group(1) if creator else "(none)",
      "| lastModifiedBy:",
      modby.group(1) if modby else "(none)")
