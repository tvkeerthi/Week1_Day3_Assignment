Daily Report Bot

A PyAutoGUI bot that opens Chrome, copies a live stock price off a webpage, opens Excel, types the data in, saves a formatted report, takes a screenshot, and closes both apps.

Why We Didn't Use Excel's "Save As"

We first tried having PyAutoGUI click through Excel's own Save As dialog. But the test machine's Excel is unlicensed, so that dialog behaved inconsistently between runs (sometimes "Save As," sometimes "Open," sometimes OneDrive) — unreliable enough that it once nearly saved over an unrelated real file. So instead, the bot types the row on-screen (to satisfy "control it like a person"), but saves the actual .xlsx file directly using openpyxl, which is reliable and safe regardless of Excel's UI state.

Files
daily_report_bot.py — the script
daily_report_YYYY-MM-DD.xlsx — generated report
excel_screenshot_YYYY-MM-DD.png — generated screenshot
