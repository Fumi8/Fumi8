# テーマの東アジア（日本語）フォントを Noto Sans JP に統一する
import sys, zipfile, re, shutil, os
src = sys.argv[1]; font = "Noto Sans JP"; tmp = src + ".tmp"
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith("ppt/theme/") and item.filename.endswith(".xml"):
            t = data.decode("utf-8")
            t = t.replace('<a:ea typeface=""/>', f'<a:ea typeface="{font}"/>')
            t = re.sub(r'<a:font script="Jpan" typeface="[^"]*"/>', f'<a:font script="Jpan" typeface="{font}"/>', t)
            data = t.encode("utf-8")
        zout.writestr(item, data)
shutil.move(tmp, src)
