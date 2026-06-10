# Anything AI

<img width="1919" height="983" alt="image" src="https://github.com/user-attachments/assets/a1a4a5fa-c839-4782-990e-4411ae5a7f98" />



### A Mixture of Experts AI Assistant for your Desktop Terminal
Built by BlueFireStudios

---

## What is Anything AI?

Anything AI is a free, open-source terminal AI assistant that routes your prompts to the best specialized AI model for the job. Instead of using one model for everything, it uses a Mixture of Experts architecture, a fast router model reads your intent and dispatches it to the most capable specialist model available.

No subscriptions. No API keys to buy. Runs on Hugging Face's free inference tier.

---

## Features

- **Intelligent Intent Routing** : Deepseek R1 Distilled 70B classifies every prompt and sends it to the right expert
- **Specialist Models** for each task:
  - `code_create` / `code_edit` → Qwen3 Coder 480B
  - `reasoning` / `math` → DeepSeek R1 Distill 70B
  - `conversation` / `roleplay` → Qwen2.5 72B
  - `agentic_task` → Qwen2.5 Coder 32B
  - `file_generate` → Qwen3 Coder 480B
- **Agentic Tasks** : open and close apps, launch websites directly from chat
- **File Generation** : create PowerPoint, Word, PDF, Excel, CSV, and text files from a single prompt
- **Image Input** : attach images via Insert key or Explorer file picker, analyzed before being sent to the model
- **Smart Memory** : keeps the last 5 exchanges raw, compresses older context automatically in the background
- **Multi-Token Fallback** : chains up to 3 Hugging Face tokens so you never hit a rate limit mid-session
- **Code Saving** : generated code is saved directly to your output folder with AI-suggested filenames
- **TUI Interface** : clean terminal UI with typewriter output, colorama coloring, and ASCII header

---

## Requirements

- Python 3.10+
- Windows (agentic features use Windows-specific app control)

### Dependencies

```
pip install huggingface_hub
pip install AppOpener
pip install python-pptx
pip install python-docx
pip install reportlab
pip install openpyxl
pip install requests
pip install Pillow
pip install pynput
pip install colorama
```

---

## Setup

**1. Download the .exe file [EXPERIMENTAL]**

Just go to the releases tab! Though be cautious as it is not ready for shipping yet.
<img width="702" height="434" alt="image" src="https://github.com/user-attachments/assets/eb142306-1188-4cb4-854e-46260c31adcc" />


**2. Set your Hugging Face token(s) as environment variables**

You need at least one free Hugging Face account. Create a token at https://huggingface.co/settings/tokens

On Windows:
```
setx HF_ACCESS_TOKEN "hf_yourtoken"
setx HF_ACCESS_TOKEN_2 "hf_yourtoken2"
setx HF_ACCESS_TOKEN_3 "hf_yourtoken3"
```

Tokens 2 and 3 are optional but recommended, the app automatically falls back to the next token if one hits its rate limit.

**3. Run**

Open the Anything AI.exe file!


<img width="414" height="363" alt="image" src="https://github.com/user-attachments/assets/484b4e2e-f2c0-4a36-9d18-0d9d994dcda9" />

---

## Usage

Just type naturally. Anything AI figures out what you need.

| What you type | What happens |
|---|---|
| `write me a snake game in python` | Qwen3 Coder generates the code, offers to save it |
| `fix the bug in the last code` | Qwen3 Coder edits with full context of previous code |
| `what is the derivative of x squared` | DeepSeek R1 solves it step by step |
| `open spotify` | AppOpener launches Spotify |
| `go to github.com` | Opens in your default browser |
| `make a presentation about black holes` | Generates a .pptx with slides and Wikimedia images |
| `what's in this image` | Press Insert to attach, model describes it |

### Keyboard Shortcuts

| Key | Action |
|---|---|
| `Insert` | Open file explorer to attach any supported file (image files supported by `PIL`, Certain text and Microsoft Office files, and code files like `.py`, `.java`, `.cpp` etc. |
| `Enter` | Send prompt (Straifghtforward, isn't it?) |

### Output Files

All generated files and saved code are stored at:
```
C:\Users\[YourUsername]\Anything AI\output\
```
And will be automatically opened on your PC
---

## Architecture

```
User Prompt
     │
     ▼
DeepSeek R1 70B (Router)
     │
     ├── code_create  ──► Qwen3 Coder 480B
     ├── code_edit    ──► Qwen3 Coder 480B
     ├── reasoning    ──► DeepSeek R1 70B
     ├── math         ──► DeepSeek R1 70B
     ├── conversation ──► Qwen2.5 72B
     ├── roleplay     ──► Qwen2.5 72B
     ├── agentic_task ──► Qwen2.5 Coder 32B
     └── file_generate──► Qwen3 Coder 480B
```

**Memory System**

Context is stored as a rolling list. The last 5 exchanges are kept raw for immediate reference. Anything older is compressed by Llama 3.1 8B in a background thread so memory never bloats without losing important details.

Long term context is saved periodically, in `Anything AI\memory\memory.txt`

**Token Fallback Chain**

Every API call tries `client1` → `client2` → `client3` in sequence. If one token is rate limited or expired, the next one takes over transparently.

---

## Limitations

- Max code output is ~2000 tokens per generation, but it depends on your Hugging Face account tier.
- Is currently Windows-only

---

## License

MIT License: free to use, modify, and distribute.

---

*Made by BlueFireStudios*


*This README was generated by Anything AI, and edited by me.*
