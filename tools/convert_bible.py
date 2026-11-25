import json, sys
sys.path.insert(0, r"D:\project\ProjectionManager\_tools\py")
from opencc import OpenCC

src = r"D:\project\ProjectionManager\Assets\bible_cuv.json"
dst = r"D:\project\ProjectionManager\Assets\bible_cuvt.json"
cc = OpenCC("s2t")

with open(src, encoding="utf-8") as f:
    data = json.load(f)

data["version"] = "和合本（繁體，由公版簡體本轉換）"
data["version_en"] = "Chinese Union Version (Traditional)"
data["source"] = "https://github.com/zjt1112112/bible-date (public domain CUV)"

for book in data["books"]:
    book["name"] = cc.convert(book["name"])
    for ch in book["chapters"]:
        for v in ch["verses"]:
            v["text"] = cc.convert(v["text"])

with open(dst, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

print("books", len(data["books"]))
print("chapters", sum(len(b["chapters"]) for b in data["books"]))
print("verses", sum(len(c["verses"]) for b in data["books"] for c in b["chapters"]))
print(data["books"][42]["name"], data["books"][42]["chapters"][2]["verses"][15]["text"])
print(data["books"][18]["name"], data["books"][18]["chapters"][22]["verses"][0]["text"])
