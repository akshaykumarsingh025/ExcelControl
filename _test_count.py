import sys, json, subprocess, os

from agent import ExcelAgent

agent = ExcelAgent(model="gemma4:31b-cloud")
b64 = agent._preprocess_image(
    "D:\\Software\\Projects\\ExcelControl\\testfiles\\Fateh Singh Gang.jpeg"
)
print(f"image: {len(b64)}b", flush=True)

prompt = (
    "Look at this table image carefully from top to bottom.\n"
    "How many total rows of data (excluding header) does this table have?\n"
    "If the S.No/row-number column is cut off, still count the data rows.\n"
    "Answer with just the number."
)

payload = {
    "model": "gemma4:31b-cloud",
    "messages": [{"role": "user", "content": prompt, "images": [b64]}],
    "stream": False,
}

with open("D:\\Software\\Projects\\ExcelControl\\_payload.json", "w") as f:
    json.dump(payload, f)

result = subprocess.run(
    ["curl", "-s", "-X", "POST", "http://localhost:11434/api/chat", "-d", "@D:\\Software\\Projects\\ExcelControl\\_payload.json"],
    capture_output=True, text=True, timeout=300,
)
os.remove("D:\\Software\\Projects\\ExcelControl\\_payload.json")

resp = json.loads(result.stdout)
print(f"Model answer: {resp['message']['content']}", flush=True)
