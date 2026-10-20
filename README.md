XTIM DAILY X THREAT INTELLIGENCE TOOL
======================================
SETUP AND DAILY RUN GUIDE
-------------------------

Purpose
-------
XTIM collects X posts using the Edge capture extension, exports the captured posts as JSON, and processes the JSON into the daily Excel workbook:

x_threat_intel_YYYY-MM-DD.xlsx

The tool also supports OCR for text inside attached images/screenshots using the bundled Tesseract-OCR folder included inside the XTIM v7 package. OCR is optional. If OCR is not available, the script still runs and processes text-only post content.


PACKAGE FILES TO KEEP
---------------------
Keep these files/folders in the XTIM folder:

1. X_TIM_script.py
   Main Python processor. Converts exported JSON into the Excel workbook.

2. X_TIM_run.bat
   Daily run helper. Double-click this after exporting the JSON.

3. XTIM_requirements.txt
   Python package list. Use this when installing requirements manually.

4. XTIM_capture_extension folder
   Edge extension folder to load through edge://extensions.

5. XTIM_README.txt
   This setup and run guide.

6. Tesseract-OCR folder
   Bundled OCR engine folder. This should already be included inside the zipped XTIM v7 package.

7. X Account Details.txt
   Login details for the X account.



OUTPUT AND LOG FILES CREATED BY THE SCRIPT
------------------------------------------
The Python script writes output and logs next to X_TIM_script.py:

- x_threat_intel_YYYY-MM-DD.xlsx
- X_TIM_logs\X_TIM_seen_posts.json
- X_TIM_logs\X_TIM_latest_run_notes.txt
- X_TIM_logs\X_TIM_run_history.log
- X_TIM_logs\ocr_cache\

Do not delete X_TIM_logs during normal use. The seen-posts file prevents duplicate rows from being added when the same posts are processed again.


FIRST-TIME SETUP
================

Step 1 - Extract the complete XTIM v7 package
---------------------------------------------
Download the zipped XTIM v7 package and extract the full folder anywhere on the PC.

After extraction, confirm the folder contains:

- X_TIM_script.py
- X_TIM_run.bat
- XTIM_requirements.txt
- XTIM_README.txt
- X Account Details
- XTIM_capture_extension\
- Tesseract-OCR\

Do not run the tool from inside the zip preview. Extract the zip first.


Step 2 - Setup X List
------------------------------------------------
Use an email ID and password to log in to X at:

https://x.com/

On X, go to the Lists tab and add as many threat intel pages, dark web monitoring and news pages, underground cybersecurity intel pages and things as you would wish to collect intel from to a list named "TIM"


Step 3 - Open Command Prompt in the XTIM folder
-----------------------------------------------
Go to the folder where X_TIM_script.py is stored.

In File Explorer, you can click the address bar, type:

cmd

Then press Enter.


Step 4 - Install Python requirements manually
---------------------------------------------
Run:

py -m pip install -r XTIM_requirements.txt

This installs the Python packages used by the processor, including openpyxl, requests, Pillow, and pytesseract.

If pip says requirements are already satisfied, that is fine.

If you ever need to confirm pytesseract is installed in the same Python used by py, run:

py -c "import pytesseract; print('pytesseract import OK')"

Use py for this project, not python, unless you know both commands point to the same Python installation.


EDGE EXTENSION SETUP
====================

Step 1 - Open the extension page
--------------------------------
Open Microsoft Edge and go to:

edge://extensions


Step 2 - Enable developer mode
------------------------------
Turn on Developer mode.

This is required so Edge allows loading the unpacked local extension folder.


Step 3 - Load the extension
---------------------------
Click:

Load unpacked

Select the extension folder named:

XTIM_capture_extension

This folder is inside the XTIM package. The folder should contain files such as:

- manifest.json
- content.js
- popup.js
- popup.html
- styles.css

Select the folder itself, not an individual file.


Step 4 - Pin the extension
--------------------------
After loading, pin the XTIM capture extension in the Edge toolbar if needed.


Step 5 - Reload the extension after updates
-------------------------------------------
If you replace or update extension files later:

1. Go back to edge://extensions.
2. Find the XTIM capture extension.
3. Click Reload.
4. Refresh any open X tabs before starting a new capture.


OCR SETUP - BUNDLED TESSERACT-OCR FOLDER
========================================

Important OCR note
------------------
The Python package pytesseract is only a Python wrapper. The actual OCR engine is tesseract.exe.

For this XTIM version, the full Tesseract-OCR folder is expected to be bundled inside the zipped XTIM v7 folder. You should not need to download or install Tesseract separately if the package was extracted correctly.

You do not need to change Windows PATH or edit environment variables.


Step 1 - Confirm Tesseract-OCR is included in the package
---------------------------------------------------------
After extracting the package, confirm this file exists:

XTIM v7\Tesseract-OCR\tesseract.exe

Also confirm this folder exists:

XTIM v7\Tesseract-OCR\tessdata\

At minimum, the tessdata folder should include:

XTIM v7\Tesseract-OCR\tessdata\eng.traineddata

If these files are present, OCR should be detected automatically by X_TIM_script.py.


Step 2 - Test direct bundled Tesseract execution
------------------------------------------------
Open Command Prompt in the XTIM folder and run:

"Tesseract-OCR\tesseract.exe" --version

Expected result:

Tesseract prints its version and library details.

If this works, the bundled OCR engine can run from the extracted folder.


Step 3 - Test XTIM OCR auto-detection
-------------------------------------
Open Command Prompt in the XTIM folder and run:

py -c "from pathlib import Path; import X_TIM_script as x; ctx=x.OCRContext(enabled=True, cache_dir=Path('X_TIM_logs')/'ocr_cache_test', base_dir=Path.cwd()); print(ctx.reason)"

Expected success result:

OCR available via ...\Tesseract-OCR\tesseract.exe ...

If you see OCR unavailable, read the message. It will list the paths that were searched.


Step 4 - Other supported locations if the bundled folder is moved
-----------------------------------------------------------------
The recommended setup is:

XTIM v7\Tesseract-OCR\tesseract.exe

The script can also detect Tesseract in these locations without PATH changes:

- XTIM v7\tools\Tesseract-OCR\tesseract.exe
- XTIM v7\ocr\Tesseract-OCR\tesseract.exe
- The parent folder next to XTIM v7, for example:

  Threat Intelligence\Tesseract-OCR\tesseract.exe

- C:\Program Files\Tesseract-OCR\tesseract.exe
- C:\Program Files (x86)\Tesseract-OCR\tesseract.exe
- C:\Users\<your user>\AppData\Local\Tesseract-OCR\tesseract.exe
- C:\Users\<your user>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe

However, for users receiving the zipped package, keep the bundled folder inside XTIM v7 unless there is a specific reason to move it.


Step 5 - If the bundled Tesseract folder is blocked by Windows
--------------------------------------------------------------
Some corporate machines may block EXE files extracted from a downloaded zip.

If Tesseract does not run, try this:

1. Right-click the original zip file.
2. Click Properties.
3. If an Unblock checkbox or button appears, select Unblock.
4. Click Apply.
5. Extract the zip again.
6. Test again:

"Tesseract-OCR\tesseract.exe" --version

If only tesseract.exe is blocked, right-click:

XTIM v7\Tesseract-OCR\tesseract.exe

Then open Properties and use Unblock if available.


Step 6 - If Tesseract is somewhere else
---------------------------------------
This should not normally be needed because Tesseract-OCR is bundled with the package.

If the folder was moved or a custom Tesseract location must be used, run the processor with an explicit path:

py -u X_TIM_script.py --tesseract-path "C:\Path\To\Tesseract-OCR\tesseract.exe"

Or if giving the folder instead of the EXE:

py -u X_TIM_script.py --tesseract-path "C:\Path\To\Tesseract-OCR"


EVERY FULL RUN WORKFLOW
=======================

Use this workflow every time you want to complete a full XTIM run.


Step 1 - Open Edge and prepare the capture extension
----------------------------------------------------
1. Open Microsoft Edge.
2. Make sure the XTIM capture extension is loaded and enabled.
3. If you recently updated extension files, go to edge://extensions and click Reload on the extension.
4. Refresh any open X tabs after reloading the extension.


Step 2 - Clear the extension capture buffer
-------------------------------------------
Open the XTIM capture extension popup and click:

Clear Capture Buffer

Do this before starting a fresh capture so the exported JSON contains only the posts from the current run.


Step 3 - Start capture
----------------------
1. Open X.
2. Go to the More tab on the left.
3. Click Lists.
4. From Your Lists, select "TIM".
5. This opens a list feed of the threat-intelligence sources and accounts.
6. Open the XTIM capture extension popup.
7. Click Start Capture.


Step 4 - Collect posts
----------------------
Scroll at a decent pace through the posts you want to capture.

The extension captures visible post text and attached image URLs. If OCR is working, the Python processor later downloads and OCRs attached images.

Tips:

- Scroll at a speed where posts still visibly load in time.
- You may click the middle mouse button and autoscroll at a medium pace.
- If the page looks stuck or incomplete, refresh the X page and continue.


Step 5 - Stop and export JSON
-----------------------------
Once most posts are beyond the time frame you want to scrape:

1. Open the XTIM capture extension popup.
2. Click Export Capture JSON.
3. Save or leave the exported JSON in one of the locations the script searches.

The script searches:

- Downloads folder
- Same folder as X_TIM_script.py
- Current working directory
- Direct subfolders of those locations

Recommended simple approach:

Leave the exported x_captured_posts*.json file in Downloads or move it into the XTIM folder.


Step 6 - Run the processor
--------------------------
Use the BAT file:

Double-click X_TIM_run.bat

Or run manually from Command Prompt in the XTIM folder:

py -u X_TIM_script.py

To process a specific JSON file manually:

py -u X_TIM_script.py "C:\path\to\x_captured_posts_file.json"

To run without OCR:

py -u X_TIM_script.py --no-ocr

To run with an explicit Tesseract path:

py -u X_TIM_script.py --tesseract-path "C:\Path\To\Tesseract-OCR\tesseract.exe"


Step 7 - Review command output
------------------------------
Check the Command Prompt output for:

- Which JSON file was selected
- OCR status
- How many records were processed
- How many new rows were added
- Workbook path
- Any warnings or errors


Step 8 - Open the generated workbook
------------------------------------
Open:

x_threat_intel_YYYY-MM-DD.xlsx

Review the sheets:

1. CRITICAL - Full Review
   Highest priority items. These should have direct company technology impact, core company-region business incident relevance, or high-confidence Microsoft identity/M365 threat relevance.

2. REVIEW - Analyst Check
   Items that need analyst review but are not automatically Critical.

3. LOW - Archive
   Items retained for monitoring/archive only.


Step 9 - Keep or archive run artifacts
--------------------------------------
Keep the workbook and logs.

Keep X_TIM_logs during normal usage because the seen-posts database prevents duplicates.

Only delete X_TIM_logs for a clean test run when you intentionally want to reprocess posts that were already seen.


POSSIBLE ERRORS AND WHAT TO DO
==============================

Error: py is not recognized
---------------------------
Meaning:
Python launcher is not available.

What to do:
Install Python for Windows and make sure the Python launcher is installed. Then reopen Command Prompt and run:

py --version


Error: No module named openpyxl, requests, PIL, or pytesseract
--------------------------------------------------------------
Meaning:
Python requirements are not installed in the Python environment used by py.

What to do:
Run this from the XTIM folder:

py -m pip install -r XTIM_requirements.txt

Then test:

py -c "import openpyxl, requests, PIL, pytesseract; print('requirements OK')"


Error: tesseract is not recognized
-----------------------------------
Meaning:
The command tesseract --version only works if Tesseract is on PATH.

What to do:
This is not required for XTIM. Use the bundled Tesseract path instead:

"Tesseract-OCR\tesseract.exe" --version

Then test XTIM detection:

py -c "from pathlib import Path; import X_TIM_script as x; ctx=x.OCRContext(enabled=True, cache_dir=Path('X_TIM_logs')/'ocr_cache_test', base_dir=Path.cwd()); print(ctx.reason)"


Error: OCR unavailable
----------------------
Meaning:
The script is running, but it cannot find or use tesseract.exe.

What to do:
Check these items:

1. Confirm this file exists:

XTIM v7\Tesseract-OCR\tesseract.exe

2. Confirm this file exists:

XTIM v7\Tesseract-OCR\tessdata\eng.traineddata

3. Test direct execution:

"Tesseract-OCR\tesseract.exe" --version

4. Test XTIM detection:

py -c "from pathlib import Path; import X_TIM_script as x; ctx=x.OCRContext(enabled=True, cache_dir=Path('X_TIM_logs')/'ocr_cache_test', base_dir=Path.cwd()); print(ctx.reason)"

5. If Windows blocked the downloaded EXE, unblock the zip or tesseract.exe from file Properties, extract again, and retest.

If needed, run without OCR:

py -u X_TIM_script.py --no-ocr


Error: No JSON found
--------------------
Meaning:
The script could not find an exported x_captured_posts*.json file.

What to do:
Move the exported JSON to one of these places:

- XTIM folder
- Downloads folder
- A direct subfolder of the XTIM folder or Downloads

Then rerun:

py -u X_TIM_script.py

Or run with the exact JSON path:

py -u X_TIM_script.py "C:\path\to\x_captured_posts_file.json"


Error: Workbook save failed or permission denied
------------------------------------------------
Meaning:
The Excel workbook is probably open, or OneDrive/Excel is locking it.

What to do:
1. Close the workbook in Excel.
2. Wait a few seconds for OneDrive sync to release the file.
3. Rerun the script.


Issue: Zero new rows added
--------------------------
Possible causes:

1. The posts were already processed before.
2. The script selected an older or wrong JSON file.
3. The JSON contains only posts outside the configured age filter.
4. The extension buffer was not cleared before capture, so the JSON was not what you expected.

What to do:

1. Confirm which JSON file the script selected in the Command Prompt output.
2. Move the correct JSON into the XTIM folder and run with the exact path.
3. For testing only, use:

py -u X_TIM_script.py --max-age-days 0

4. For a clean reprocessing test only, delete X_TIM_logs or X_TIM_logs\X_TIM_seen_posts.json. Do not do this during normal daily usage unless you intentionally want duplicates/reprocessing.


Issue: Extension captures nothing
---------------------------------
Possible causes:

1. Extension was not reloaded after file updates.
2. X tab was opened before loading/reloading the extension.
3. Capture was not started.
4. The page content did not fully load.

What to do:

1. Go to edge://extensions.
2. Click Reload on the XTIM capture extension.
3. Refresh the X page.
4. Open extension popup and click Start Capture.
5. Scroll slowly and export again.


Issue: Excel has too many false positives
-----------------------------------------
What to do:
Use the v7.2 precision script or later. v7.2 reduces false positives by requiring company names and generic technology terms to appear in direct threat context before they promote a record.

If a repeated false positive still appears, adjust the general matching logic rather than hardcoding one title or one post.


Issue: Excel is missing something that should be Review
-------------------------------------------------------
What to do:
Check whether the post has direct company technology, company brand, company region, M365 identity, payment/ecommerce skimming, RMM phishing, FortiGate access sale, logistics/vendor exposure, or critical web-app CVE context.

If the missed item represents a repeatable pattern, add a general special relevance pattern rather than a one-off rule.


USEFUL MANUAL COMMANDS
======================

Install requirements:

py -m pip install -r XTIM_requirements.txt

Check Python used by py:

py -c "import sys; print(sys.executable)"

Check Python package import:

py -c "import pytesseract; print('pytesseract import OK')"

Check bundled Tesseract directly:

"Tesseract-OCR\tesseract.exe" --version

Check OCR auto-detection:

py -c "from pathlib import Path; import X_TIM_script as x; ctx=x.OCRContext(enabled=True, cache_dir=Path('X_TIM_logs')/'ocr_cache_test', base_dir=Path.cwd()); print(ctx.reason)"

Check script syntax:

py -m py_compile X_TIM_script.py

Run default processing:

py -u X_TIM_script.py

Run without OCR:

py -u X_TIM_script.py --no-ocr

Run a specific JSON:

py -u X_TIM_script.py "C:\path\to\x_captured_posts_file.json"

Run with explicit Tesseract path:

py -u X_TIM_script.py --tesseract-path "C:\Path\To\Tesseract-OCR\tesseract.exe"

Disable age filtering for testing:

py -u X_TIM_script.py --max-age-days 0

Use a specific workbook date:

py -u X_TIM_script.py --date YYYY-MM-DD
