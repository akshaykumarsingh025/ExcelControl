import sys, json, subprocess, os, re

from agent import ExcelAgent

agent = ExcelAgent(model="gemma4:31b-cloud")
b64 = agent._preprocess_image(
    "D:\\Software\\Projects\\ExcelControl\\testfiles\\Fateh Singh Gang.jpeg"
)
print(f"image: {len(b64)}b", flush=True)

# Use the same prompt as agent.ask_with_image
prompt = (
    "Extract EVERY row from this table and write it to the sheet starting at A1.\n\n"
    "STEP 1 — COUNT:\n"
    "Count the total number of DATA rows (not counting the header). "
    "IMPORTANT: The row-number column (S.No, #, etc.) may be partially cut off "
    "in the image — ignore it. Count rows by looking at the actual data entries "
    "(names, dates, bank names, etc.), not the index numbers. "
    "If you see data in a row but no row number, still include that row.\n\n"
    "STEP 2 — LIST ALL DATA:\n"
    "Write: data = [\n"
    '    ["header1", "header2", ...],\n'
    '    ["row1col0", "row1col1", ...],\n'
    "    ... every single row ...\n"
    "]\n"
    "Numbers as int/float, text as strings, empty cells as None. "
    "If a row has no S.No number but has data, still include it.\n\n"
    "STEP 3 — XLWINGS CODE:\n"
    'ws.range("A1").value = data\n'
    'ws.range("A1").expand("right").font.bold = True\n'
    "ws.autofit()\n\n"
    "CRITICAL: Look at the ENTIRE image. Do not stop at the last visible row number. "
    "Extract absolutely every row that has data — missing even one row is a failure."
)

payload = {
    "model": "gemma4:31b-cloud",
    "messages": [{"role": "user", "content": prompt, "images": [b64]}],
    "stream": False,
}

tmpfile = "D:\\Software\\Projects\\ExcelControl\\_payload.json"
with open(tmpfile, "w") as f:
    json.dump(payload, f)

print("sending...", flush=True)
result = subprocess.run(
    ["curl", "-s", "-X", "POST", "http://localhost:11434/api/chat", "-d", f"@{tmpfile}"],
    capture_output=True, text=True, timeout=600,
)
os.remove(tmpfile)

resp = json.loads(result.stdout)
extracted = resp["message"]["content"]
clean = re.sub(r"```python|```", "", extracted).strip()

# Count data rows from the "data = [" list
data_match = re.search(r'data\s*=\s*\[(.*?)\]', clean, re.DOTALL)
if data_match:
    inner = data_match.group(1)
    row_count = inner.count('"],') + 1 if '"],' in inner else 0
    # subtract header row
    data_rows = max(0, row_count - 1)
    print(f"\nData rows extracted: {data_rows}", flush=True)
else:
    print("\nCould not parse data rows", flush=True)

print("\n" + "=" * 70)
# Print only the data section, not full code
idx = clean.find("data = [")
if idx >= 0:
    end = clean.find("]", idx) + 1
    print(clean[idx:end])
print("=" * 70)
