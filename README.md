# WhatsApp Web Automation Bot

This project sends messages through WhatsApp Web using Selenium and Chrome on Windows.

Use it only for contacts who gave permission to receive your messages. Unwanted bulk messaging can get your WhatsApp account restricted.

## Folder Structure

```text
whatsapp-bot/
├── bot.py
├── contacts.xlsx
├── chromedriver.exe        # optional with Selenium 4.6+
├── logs.txt
├── image.jpg              # optional attachment example
├── report.xlsx            # generated after a run
└── chrome-data/           # generated automatically for saved login
```

## Setup

1. Open PowerShell in this folder:

   ```powershell
   cd "C:\Users\aspdi\OneDrive\Desktop\Whatsapp autosender\Wharsapp sender script\whatsapp-bot"
   ```

2. Create and activate a virtual environment:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

3. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

4. Put your contacts in `contacts.xlsx`.

   Required column:

   | number |
   | ------ |
   | 919876543210 |

   Use country code and do not use `+`.

5. Optional: add `chromedriver.exe` to this folder.

   Selenium 4.6+ can usually manage ChromeDriver automatically. If Chrome does not open, download the ChromeDriver version matching your Chrome browser and place it beside `bot.py`.

6. You can still edit defaults in the config section at the top of `bot.py`,
   but the bot now asks for the most important options every time it starts:

   - minimum delay
   - maximum delay
   - daily sending limit
   - whether to skip numbers already sent in `logs.txt` (default is no)
   - message templates, up to 5
   - whether to send media/document
   - attachment file path

   Supported attachment types: `jpg`, `png`, `pdf`.

## Multi-line Templates

For templates with blank lines, emojis, and WhatsApp formatting, edit `templates.txt`.
Do not type long multi-line templates directly into PowerShell or CMD.

Each template should be separated with this line:

```text
---TEMPLATE---
```

You can add a label line at the top of each template so you can identify it.
The bot removes this label before sending the WhatsApp message.

Supported label formats:

```text
# Template 1 - Future Dice main offer
[Template 2 - Website offer]
Template 3: Ads offer
// Template 4 - Follow up
```

Example:

```text
# Template 1 - Future Dice main offer
Hi 👋

I came across your business and honestly felt it has **strong potential** to grow online 🚀

We at *Future Dice* help businesses get **more customers & sales** through:

✔ *Professional Websites*
✔ *Google & Meta Ads*
✔ *Social Media Marketing*
✔ *Lead Generation*

Would you like some **FREE suggestions** to grow your business online? 📈

👉 Reply *YES* and let’s discuss how we can help your business grow.

📞 *+91 7048907028*
📞 *+91 9310824993*

---TEMPLATE---
# Template 2 - Short version
Hello 👋

Would you like free suggestions to grow your business online?
```

When the bot starts, it will ask whether to load templates from `templates.txt`.
It will also ask how many loaded templates you want to use. If you added only
3 templates, the default will be 3 and the bot will randomize only between
those 3 templates.

## Run

```powershell
python bot.py
```

Startup flow:

1. The terminal asks for delay, templates, and media/document options.
2. Chrome opens WhatsApp Web.
3. Scan the QR code the first time only.
4. The script detects your logged-in account.
5. It waits for this confirmation:

   ```text
   [INFO] Press ENTER to start sending messages...
   ```

6. Press ENTER when you are ready.

## Output

The bot writes results to:

- `logs.txt`
- `report.xlsx`

The Excel report includes:

- Number
- Status
- Time
- Template Used
- Error Reason
- Summary totals for sent, failed, and skipped

## Notes

- Keep Chrome visible while the bot is running.
- Do not use the keyboard or mouse inside the WhatsApp tab while messages are being sent.
- Default delay prompts are now minimum `5` seconds and maximum `120` seconds.
- The bot pastes messages from the Windows clipboard so emojis and multi-line templates work better.
- If you want resume behavior, answer `y` when asked whether to skip numbers already marked `Sent` in `logs.txt`.
