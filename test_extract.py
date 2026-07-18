from agent import ExcelAgent
a = ExcelAgent()
a.set_model("gemma4:31b-cloud")
result = a.ask_with_image("testfiles/Fateh Singh Gang.jpeg")
print(result)
