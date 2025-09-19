# caption.py
import sys, json, os, re
from PIL import Image
import torch
from transformers import BlipProcessor, BlipForConditionalGeneration

MODEL_DIR = os.environ.get("BLIP_MODEL_DIR", "blip-image-captioning-base")

# ---- 追加：英文キャプションをタグ列に整形する軽量関数 -----------------
_STOP = {
    "a","an","the","this","that","these","those",
    "with","and","or","of","in","on","at","by","for","from","to","into","over","under","near",
    "is","are","was","were","be","being","been","there","it","its","his","her","their","our","your",
    "as","about","around","between","behind","before","after","without","within",
}
_EXCEPT_S = {"glass","gas","grass","news","lens","dress"}  # 単純な s 除去の例外
# ----------------------------------------------------------------------

def caption_to_tags(text):
    # lower & 句読点→空白
    s = text.lower()
    s = re.sub(r"[^a-z0-9\s\-]", " ", s)
    # つなぎ言葉をカンマに寄せる（and/with/of 等）
    for k in (" with ", " and ", " of ", " featuring ", " having "):
        s = s.replace(k, ", ")
    # 分割
    raw = [w for w in re.split(r"[,\s]+", s) if w]
    # ストップワード除去＋簡易単数化＋重複排除（出現順保持）
    seen, out = set(), []
    for w in raw:
        if w in _STOP or len(w) <= 1: 
            continue
        if w.endswith("s") and w[:-1] not in _EXCEPT_S and len(w) > 3:
            w = w[:-1]                   # dogs -> dog, robots -> robot
        if w not in seen:
            seen.add(w); out.append(w)

    # 簡単なビッグラム（色＋背景など）
    COLORS = {"white","black","blue","red","green","yellow","purple","pink","orange","brown","gray","grey"}
    merged = []
    i = 0
    while i < len(out):
        if i+1 < len(out) and out[i] in COLORS and out[i+1] in {"background","sky","hair","eyes"}:
            merged.append(out[i] + " " + out[i+1])   # white background 等
            i += 2
        else:
            merged.append(out[i]); i += 1

    # ほどよい長さに（長すぎ回避の調整）
    if len(merged) > 24:
        merged = merged[:24]

    return ", ".join(merged)
# ----------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print(json.dumps({"error":"no input"})); return
    path = sys.argv[1]
    image = Image.open(path).convert("RGB")

    torch.set_num_threads(max(1, os.cpu_count()//2))

    processor = BlipProcessor.from_pretrained(MODEL_DIR, local_files_only=True)
    model     = BlipForConditionalGeneration.from_pretrained(MODEL_DIR, local_files_only=True)

    inputs = processor(image, return_tensors="pt")
    out = model.generate(**inputs, max_new_tokens=40)
    cap = processor.decode(out[0], skip_special_tokens=True).strip()

    tags = caption_to_tags(cap)
    print(json.dumps({"caption": cap, "tags": tags}, ensure_ascii=False))

if __name__ == "__main__":
    main()
