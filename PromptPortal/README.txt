ANALYSIS PROMPTS PORTAL
=======================

Open index.html in any modern browser (Chrome, Edge, Firefox).

HOW TO USE
----------
- You will see 8 beautiful cards: 4 for India + 4 for USA.
- Click any card to open the full prompt.
- The portal **always injects the current system date** at the top of the prompt:
    • India prompts → "Current date (IST): YYYY-MM-DD"   (using Asia/Kolkata)
    • USA prompts  → "Current date (CST): YYYY-MM-DD"   (using America/Chicago)
- Click "Copy Prompt to Clipboard" — the text you copy already contains the correct date line + the original prompt.
- Paste directly into your LLM. It will know the exact "today" in the market's timezone.

This happens live in the browser (no need to re-generate for the date itself). Re-generate only when you change the base .docx content.

UPDATING PROMPTS
----------------
1. Edit the original Word documents inside the Analysis folder (Analysis/India or Analysis/USA).
2. Run the generator:
     - Go into the PromptPortal folder (keep it as a sibling to the Analysis folder)
     - Right-click generate-prompt-portal.ps1 → "Run with PowerShell"
     - Or open PowerShell in PromptPortal and run:  .\generate-prompt-portal.ps1
3. Refresh / re-open index.html

The whole "Analysis" folder (containing your prompts) can be moved or renamed frequently.
Just keep a "PromptPortal" folder next to it and run the generator from inside PromptPortal.

The prompts are embedded at generation time (single self-contained HTML file, works offline, no server needed).

FILES
-----
- index.html                 → The portal you open
- generate-prompt-portal.ps1 → Run this from the PromptPortal folder to refresh

Created: 2026-06-14
Prompts source: ../Analysis (relative to PromptPortal)
