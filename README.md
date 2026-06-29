# 📲 WhatsApp Web Automation Bot

> Automate WhatsApp messages through WhatsApp Web using **Python + Selenium + Chrome** — with multi-template rotation, media attachments, smart delay control, resume support, and full Excel reporting.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?style=for-the-badge&logo=python&logoColor=white)
![Selenium](https://img.shields.io/badge/Selenium-4.20%2B-43B02A?style=for-the-badge&logo=selenium&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Active-brightgreen?style=for-the-badge)

---

## ⚠️ Disclaimer

> **Use responsibly and ethically.**
> Only send messages to contacts who have **explicitly given permission** to receive them.
> Unsolicited bulk messaging violates [WhatsApp's Terms of Service](https://www.whatsapp.com/legal/terms-of-service) and can result in your account being **temporarily restricted or permanently banned**.
> The developer holds **no responsibility** for any misuse of this tool.

---

## 📋 Table of Contents

- [✨ Features](#-features)
- [📁 Folder Structure](#-folder-structure)
- [🔧 Prerequisites](#-prerequisites)
- [🚀 Installation](#-installation)
- [⚙️ Configuration](#️-configuration)
- [🖥️ How to Use](#️-how-to-use)
- [📝 Template Guide](#-template-guide)
- [📊 Output & Reports](#-output--reports)
- [📦 Dependencies](#-dependencies)
- [📌 Notes & Tips](#-notes--tips)
- [🛡️ License](#️-license)

---

## ✨ Features

| Feature | Description |
|--------|-------------|
| 🔀 **Multi-Template Rotation** | Define up to **5 message templates**; the bot randomly picks one per contact to avoid detection |
| 📎 **Media & Document Attachments** | Attach `jpg`, `png`, or `pdf` files alongside your messages |
| ⏱️ **Smart Human-like Delays** | Set custom **min/max delay** between sends to mimic natural behavior |
| 📊 **Excel Contact Import** | Load all contacts directly from `contacts.xlsx` with zero formatting hassle |
| 📝 **Persistent Logging** | Every action is logged to `logs.txt` in real time |
| 📈 **Auto Excel Report** | A detailed `report.xlsx` is generated after every run |
| 🔁 **Resume / Skip-Sent Support** | Pick up where you left off — skip numbers already marked `Sent` |
| 📅 **Daily Sending Limit** | Cap how many messages are sent per session |
| 🔐 **Persistent Login** | Chrome profile saves your session — scan QR code **only once** |
| 📋 **Clipboard Paste Engine** | Messages are pasted via clipboard so **emojis & multi-line formatting** are preserved perfectly |
| 🏷️ **Labeled Templates** | Identify and manage templates using labels inside `templates.txt` |
| 🛑 **Graceful Error Handling** | Failed numbers are logged with the reason and skipped without crashing |

---

## 📁 Folder Structure

```text
whatsapp-bot/
│
├── bot.py                  # 🤖 Main bot script
├── contacts.xlsx           # 📇 Your contact list (required)
├── templates.txt           # 📝 Multi-line message templates
├── requirements.txt        # 📦 Python dependencies
│
├── chromedriver.exe        # ⚙️ Optional — Selenium 4.6+ handles this automatically
├── image.jpg               # 📎 Optional attachment example
│
├── logs.txt                # 📋 Auto-generated run log
├── report.xlsx             # 📊 Auto-generated Excel report after each run
└── chrome-data/            # 🔐 Auto-generated Chrome profile (saved WhatsApp login)
```

---

## 🔧 Prerequisites

Make sure you have the following installed before getting started:

- ✅ **Windows 10 or 11**
- ✅ **Python 3.8+** → [Download Python](https://www.python.org/downloads/)
- ✅ **Google Chrome** (latest version) → [Download Chrome](https://www.google.com/chrome/)
- ✅ **An active WhatsApp account** linked to your phone

> **💡 ChromeDriver Note:**
> Selenium 4.6+ automatically manages ChromeDriver. If Chrome fails to open, manually download the ChromeDriver matching your Chrome version from [chromedriver.chromium.org](https://chromedriver.chromium.org/) and place `chromedriver.exe` inside the project folder.

---

## 🚀 Installation

### Step 1 — Clone or Download the Repository

```powershell
git clone https://github.com/yourusername/whatsapp-bot.git
cd whatsapp-bot
```

> Or click **Code → Download ZIP**, extract it, and open that folder in PowerShell.

---

### Step 2 — Navigate to the Project Folder

```powershell
cd "C:\Users\YourName\Desktop\whatsapp-bot"
```

---

### Step 3 — Create a Virtual Environment

```powershell
python -m venv .venv
```

---

### Step 4 — Activate the Virtual Environment

```powershell
.\.venv\Scripts\Activate.ps1
```

> ⚠️ If you see an **execution policy error**, run this command first, then retry Step 4:
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

You should see `(.venv)` appear at the start of your terminal prompt, confirming it's active.

---

### Step 5 — Install All Dependencies

```powershell
pip install -r requirements.txt
```

---

## ⚙️ Configuration

### 📇 Setting Up Your Contact List

Open or create `contacts.xlsx`. Add a single column with the header `number`:

| number |
|--------|
| 919876543210 |
| 447911123456 |
| 12025550194 |

> **Important formatting rules:**
> - ✅ Include the **full country code** (e.g., `91` for India, `44` for UK, `1` for USA)
> - ❌ Do **NOT** include `+`, spaces, dashes, or brackets
> - Example: `+91 98765 43210` → should be entered as `919876543210`

---

### 🎛️ Runtime Options

Every time the bot starts, it will interactively ask you for these settings:

| Prompt | Description | Default |
|--------|-------------|---------|
| Minimum Delay | Shortest wait between messages (seconds) | `5` |
| Maximum Delay | Longest wait between messages (seconds) | `120` |
| Daily Sending Limit | Max messages to send in this session | *(you choose)* |
| Skip Already-Sent Numbers | Skip contacts marked `Sent` in `logs.txt` | `No` |
| Load from `templates.txt`? | Use the file-based template system | *(you choose)* |
| Number of Templates to Use | How many loaded templates to rotate between | All loaded |
| Send Media/Document? | Attach a file to every message | *(you choose)* |
| Attachment File Path | Full path to the file to attach | *(you enter)* |

---

## 🖥️ How to Use

### ▶️ Step 1 — Run the Bot

```powershell
python bot.py
```

---

### 💬 Step 2 — Answer the Startup Prompts

The terminal will walk you through all settings interactively. Answer each question and press **Enter**.

---

### 📱 Step 3 — Scan the QR Code *(First Time Only)*

Chrome will launch and open WhatsApp Web automatically.

1. Open **WhatsApp** on your phone
2. Tap the **Menu icon (⋮)** → **Linked Devices**
3. Tap **Link a Device**
4. Point your camera at the QR code on screen

> ✅ Your session is saved in `chrome-data/`. You will **not** need to scan again on future runs unless you get logged out.

---

### ✅ Step 4 — Wait for Account Confirmation

Once logged in, the terminal will display:

```text
[INFO] Logged in as: +91 XXXXXXXXXX
[INFO] Press ENTER to start sending messages...
```

---

### 🚀 Step 5 — Start Sending

Press **ENTER** and the bot will begin. You'll see live status updates in the terminal:

```text
[INFO] Sending to: 919876543210 — Template 2
[INFO] ✅ Sent successfully!
[INFO] Waiting 47 seconds before next message...
```

---

## 📝 Template Guide

For **short, single-line messages** you can type them directly in the terminal when prompted.

For **multi-line messages with emojis, bold/italic formatting, or blank lines**, always use `templates.txt` — typing these directly into PowerShell will break the formatting.

---

### 📄 Format of `templates.txt`

Separate each template with this exact divider on its own line:

```
---TEMPLATE---
```

Each template can optionally have a **label on the first line** (the bot strips it before sending). Supported label formats:

```text
# Template 1 - Main Offer
[Template 2 - Website Offer]
Template 3: Ad Campaign
// Template 4 - Follow Up
```

---

### 📋 Full Example `templates.txt`

```
# Template 1 - Future Dice main offer
Hi 👋

I came across your business and honestly felt it has *strong potential* to grow online 🚀

We at *Future Dice* help businesses get *more customers & sales* through:

✔ *Professional Websites*
✔ *Google & Meta Ads*
✔ *Social Media Marketing*
✔ *Lead Generation*

Would you like some *FREE suggestions* to grow your business online? 📈

👉 Reply *YES* and let's discuss how we can help your business grow.

📞 *+91 7048907028*
📞 *+91 9310824993*

---TEMPLATE---

# Template 2 - Short version
Hello 👋

Would you like free suggestions to grow your business online?

Reply *YES* and I'll share some quick wins for your brand! 🚀

📞 *+91 7048907028*

---TEMPLATE---

# Template 3 - Follow Up
Hi again 👋

Just following up on my earlier message about growing your business online.

We've already helped several businesses in your area get more leads and sales.

Interested? Reply *YES* and I'll send you a free audit 📊
```

---

### ✏️ WhatsApp Formatting Reference

| Style | Syntax | Result |
|-------|--------|--------|
| **Bold** | `*your text*` | **your text** |
| *Italic* | `_your text_` | *your text* |
| ~~Strikethrough~~ | `~your text~` | ~~your text~~ |
| `Monospace` | ` ```your text``` ` | `your text` |

---

## 📊 Output & Reports

After every run the bot generates two output files:

### 📋 `logs.txt`

A plain-text log capturing every action in real time:

```text
[2025-01-15 10:23:01] [SENT]    919876543210 — Template 2
[2025-01-15 10:25:44] [FAILED]  917890123456 — Number not on WhatsApp
[2025-01-15 10:28:12] [SKIPPED] 919123456789 — Already sent previously
[2025-01-15 10:31:05] [SENT]    919001234567 — Template 1
```

---

### 📊 `report.xlsx`

A structured Excel report generated at the end of each session:

| Column | Description |
|--------|-------------|
| **Number** | The contact number |
| **Status** | `Sent`, `Failed`, or `Skipped` |
| **Time** | Exact timestamp of the attempt |
| **Template Used** | Which template was selected for this contact |
| **Error Reason** | Why it failed (if applicable) |

The bottom of the report includes **summary totals**:

```
✅ Total Sent:     142
❌ Total Failed:    8
⏭️ Total Skipped:  12
```

---

## 📦 Dependencies

All dependencies are listed in `requirements.txt` and installed with a single command.

| Package | Min Version | Purpose |
|---------|-------------|---------|
| `selenium` | `4.20.0` | Browser automation to control Chrome |
| `pandas` | `2.0.0` | Reading `contacts.xlsx` |
| `openpyxl` | `3.1.0` | Writing the Excel `report.xlsx` |
| `pyautogui` | `0.9.54` | Clipboard interaction for emoji-safe pasting |

```powershell
pip install -r requirements.txt
```

---

## 📌 Notes & Tips

- 🖥️ **Keep Chrome visible** — Do not minimize the Chrome window while the bot is running
- 🖱️ **Hands off the WhatsApp tab** — Do not click or type in the browser tab during sending
- 🌐 **Stable internet required** — An unstable connection may cause send failures
- 🔄 **Resuming a session** — Answer `y` to "skip already-sent numbers" to continue from where you stopped
- 📎 **Supported attachment types** — `.jpg`, `.png`, `.pdf` only
- ⏱️ **Recommended delay** — Use at least `15–30s` minimum and `90–120s` maximum to reduce ban risk
- 🔒 **Do not share `chrome-data/`** — It contains your WhatsApp session and should be kept private
- 🛑 **To stop the bot** — Press `Ctrl + C` in the terminal at any time

---

## 🛡️ License

This project is licensed under the **MIT License**.
You are free to use, modify, and distribute it with attribution.
See the [LICENSE](LICENSE) file for full details.

---

## 🙏 Acknowledgements

- [Selenium](https://www.selenium.dev/) — Powerful browser automation framework
- [WhatsApp Web](https://web.whatsapp.com/) — The platform this bot interfaces with
- [OpenPyXL](https://openpyxl.readthedocs.io/) — Excel file creation and editing
- [Pandas](https://pandas.pydata.org/) — Fast and flexible data handling
- [PyAutoGUI](https://pyautogui.readthedocs.io/) — Clipboard and keyboard automation

---

<div align="center">

Made with ❤️ | For legitimate business outreach only

⭐ **If this project helped you, consider giving it a star!** ⭐

</div>

---

> ⚠️ **Final Reminder:** This tool is designed for **legitimate, consent-based outreach only**.
> Always obtain explicit permission before messaging contacts.
> Misuse may violate WhatsApp's Terms of Service and result in account suspension.
