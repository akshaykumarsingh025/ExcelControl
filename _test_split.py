import sys, json, subprocess, os
from PIL import Image
from agent import ExcelAgent
import base64, io

agent = ExcelAgent(model="gemma4:31b-cloud")

# Load original and split
img = Image.open("D:\\Software\\Projects\\ExcelControl\\testfiles\\Fateh Singh Gang.jpeg")
w, h = img.size
print(f"Original: {w}x{h}", flush=True)

# Split into top and bottom halves with overlap
mid = h // 2
overlap = 60
top_half = img.crop((0, 0, w, mid + overlap))
bot_half = img.crop((0, mid - overlap, w, h))

top_half.save("D:\\Software\\Projects\\ExcelControl\\_top.png")
bot_half.save("D:\\Software\\Projects\\ExcelControl\\_bot.png")

for label, path in [("TOP HALF", "_top.png"), ("BOTTOM HALF", "_bot.png")]:
    # Preprocess the same way as ExcelAgent
    from PIL import Image as PILImage
    import PIL.ImageEnhance as IE, PIL.ImageFilter as IF
    
    crop = PILImage.open(path)
    max_dim = 4096
    if max(crop.size) > max_dim:
        crop.thumbnail((max_dim, max_dim), PILImage.LANCZOS)
    if max(crop.size) < 3072:
        s = 3072 / max(crop.size)
        crop = crop.resize((int(crop.width*s), int(crop.height*s)), PILImage.LANCZOS)
    
    crop = crop.convert("L")
    crop = crop.filter(IF.MedianFilter(3))
    crop = IE.Contrast(crop).enhance(2.0)
    crop = IE.Sharpness(crop).enhance(2.5)
    crop = crop.convert("RGB")
    
    buf = io.BytesIO()
    crop.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    
    print(f"\n{label} ({crop.size}) -> {len(b64)}b", flush=True)
    
    prompt = "How many data rows (excluding header) are in this table? Answer with just a number."
    payload = {
        "model": "gemma4:31b-cloud",
        "messages": [{"role": "user", "content": prompt, "images": [b64]}],
        "stream": False,
    }
    with open("D:\\Software\\Projects\\ExcelControl\\_payload.json", "w") as f:
        json.dump(payload, f)
    
    try:
        result = subprocess.run(
            ["curl", "-s", "-X", "POST", "http://localhost:11434/api/chat", "-d", "@D:\\Software\\Projects\\ExcelControl\\_payload.json"],
            capture_output=True, text=True, timeout=180,
        )
        resp = json.loads(result.stdout)
        print(f"  -> Rows: {resp['message']['content']}", flush=True)
    except Exception as e:
        print(f"  -> Error: {e}", flush=True)

os.remove("D:\\Software\\Projects\\ExcelControl\\_top.png")
os.remove("D:\\Software\\Projects\\ExcelControl\\_bot.png")
os.remove("D:\\Software\\Projects\\ExcelControl\\_payload.json")
print("\nDone", flush=True)
