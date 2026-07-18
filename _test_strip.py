import sys
sys.path.insert(0, "D:\\Software\\Projects\\ExcelControl")

from agent import ExcelAgent

agent = ExcelAgent(model="gemma4:31b-cloud")
print("Starting strip-based extraction...", flush=True)
code = agent.ask_with_image(
    "D:\\Software\\Projects\\ExcelControl\\testfiles\\Fateh Singh Gang.jpeg"
)

# Count rows from the data list
import re, json
match = re.search(r"data\s*=\s*(\[[\s\S]*?\])\n", code)
if match:
    try:
        data = eval(match.group(1))
        print(f"\nTotal rows (incl header): {len(data)}", flush=True)
        print(f"Data rows: {len(data) - 1}", flush=True)
    except:
        pass

print("\n" + "=" * 70)
print(code[:5000])
if len(code) > 5000:
    print(f"\n... ({len(code) - 5000} more chars)")
print("=" * 70)
