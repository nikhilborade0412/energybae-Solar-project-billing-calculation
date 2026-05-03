import google.generativeai as genai

# 🔑 Set your API key
genai.configure(api_key="AIzaSyCJkd1psilq5jw49i1tuPQvbwWoNkYKRvg")

# 📋 List all models
models = genai.list_models()

print("Available Models:\n")

for model in models:
    # Show only models that support text generation
    if "generateContent" in model.supported_generation_methods:
        print(f"👉 {model.name}")