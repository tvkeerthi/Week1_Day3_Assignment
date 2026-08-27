WhatsApp Web Automation Bot

A Playwright bot that logs into WhatsApp Web, sends personalized messages to contacts, and saves the results as reports.

What It Does

Reads contacts from contacts.xlsx → opens WhatsApp Web → searches and messages each contact → extracts their last 3 messages → saves whatsapp_report_<date>.json and .xlsx.

Notes
Selectors were found by inspecting the live page in DevTools (WhatsApp's class names change often, so we target stable attributes like aria-label/data-testid instead).
Before sending, the bot verifies the opened chat actually matches the intended contact — stops instead of risking a message to the wrong person.
Random delays between actions to avoid looking robotic.
Setup & Run
pip install playwright openpyxl
playwright install
python playwright_assign.py

Scan the QR code on first run only. Only message contacts who've agreed to hear from you.
