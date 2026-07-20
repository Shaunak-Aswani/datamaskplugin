# DataMask

A Claude plugin for anonymising client-identifying information (entity
names, individual names, addresses, emails, phone numbers, registration
numbers) in Word, Excel, PowerPoint, and PDF documents before sharing
them externally, building sample/training libraries, or doing internal
review.

## What's included

- **`skills/datamask/SKILL.md`** - the skill itself. Once installed,
  just describe what you want in chat (e.g. "anonymise this SOW for
  Novotech, call the placeholder NeoPath") and Claude will follow the
  scan -> human review -> apply workflow described in the skill, or
  invoke it directly with `/datamask:datamask`.
- **`skills/datamask/scripts/`** - the underlying Python
  (`datamask_core.py`, `datamask_cli.py`) that does the actual
  detection/masking, plus a standalone Streamlit web app
  (`datamask_webapp.py` + `datamask_webapp_logic.py`) for anyone who'd
  rather review and apply masks through a browser UI instead of chat.

## Installing this plugin

If this repo is hosted on GitHub (or another git remote), add it as a
marketplace and install from it:

```
/plugin marketplace add <your-repo-url>
/plugin install datamask@<marketplace-name>
```

In Claude.ai / Claude Cowork: Customize menu -> Plugins tab -> "+" ->
Add marketplace, then install `datamask` from the list. (Marketplace
installs require a Pro or Team plan.)

## Using the standalone web app instead

If someone would rather not use the chat/skill workflow at all:

```bash
cd skills/datamask/scripts
python3 -m venv venv
source venv/bin/activate
pip install streamlit spacy python-docx openpyxl python-pptx pymupdf pandas pillow
python -m spacy download en_core_web_sm
streamlit run datamask_webapp.py
```

## Never share client documents through an untrusted third-party plugin

This plugin runs entirely with your own local Python code - no document
content is sent anywhere outside the scan -> review -> apply steps you
run yourself. If distributing this further within your organisation,
make sure whoever installs it understands the same principle applies to
anything else they install from a marketplace: never install a
plugin that can see confidential documents unless you trust its source.
