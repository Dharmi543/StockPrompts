ANALYSIS PROMPTS PORTAL
=======================

Open index.html in any modern browser (Chrome, Edge, Firefox).

HOW TO USE
----------
- You will see 10 beautiful cards: 5 for India + 5 for USA.
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

VERIFICATION TIP
----------------
Every time you run generate-prompt-portal.ps1, it now prints a line for each of the 10 prompts:
  [OK]   Indian Market News             1570 chars  | Indian Stock Market Analysis Prompt: "Please identify...

- Look for [OK] on all 10 lines.
- The preview after the | should look like the beginning of your prompt in the .docx.
- If you see [ERROR] or very short length, something went wrong with that specific document — check the .docx or share the output here.

The extraction was upgraded (June 2026) to use real XML parsing instead of fragile regex, so it should now reliably pull the full prompt text from your Word documents.

TROUBLESHOOTING: "FILE IN USE" / "being used by another process"
----------------------------------------------------------------
If you see an extraction error like:

  [FILE IN USE] The .docx file is currently locked...

This almost always means one of these:
- The .docx is open in Microsoft Word
- OneDrive is actively syncing the file
- Another program has a lock on it

How to fix:
1. Completely close the document in Word (File → Close, or exit Word).
2. Wait 3–5 seconds for OneDrive to finish syncing.
3. Re-run generate-prompt-portal.ps1 from the PromptPortal folder.
4. Refresh the portal page.

The generator now has retries + much clearer error messages for this exact situation.
