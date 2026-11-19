# Screen Data Logger – Updated OCR Retry Edition

## What changed
- Parameters are read **sequentially**, one by one, during each logging cycle.
- Each parameter has a configurable OCR retry count (0–10 retries after the first attempt); default is **5 retries**, for up to 6 total attempts.
- If OCR returns blank/unreadable text or raises an error, the application retries that parameter using different image preprocessing and Tesseract page segmentation modes.
- A successful retry is recorded with its attempt number.
- If all attempts fail, the Excel value cell is left blank and the status column records `OCR FAILED (5 attempts)`. The previous reading is never reused.
- The first sample is taken immediately after Start; subsequent samples follow the configured logging interval.
- Existing parameter configuration is retained in the user's home folder.

## Install
1. Install Python 3.10 or newer and tick **Add Python to PATH** during installation.
2. Install Tesseract OCR for Windows. Common installation path:
   `C:\Program Files\Tesseract-OCR\tesseract.exe`
   If installed elsewhere, set the `TESSERACT_CMD` environment variable to the full `tesseract.exe` path.
3. Extract this ZIP to a normal folder.
4. Double-click `Run_Screen_Data_Logger.bat`. It installs the Python packages and launches the application.

## Use
1. Add a parameter for each displayed SCADA value.
2. Select a parameter row and click **Select rectangle for selected row**; drag around that value only.
3. Repeat for every parameter. You can enable/disable rows and edit parameter name, unit, and decimals in the table.
4. Set the logging interval, retry attempts (default 5), and Excel output path.
5. Use **Test OCR (with retries)** for every selected region before starting.
6. Click **Start logging**. Excel is saved after every complete logging cycle.

## Important behaviour
- Each parameter is captured and OCR-processed one at a time, not in parallel.
- Retries are per parameter. A failed parameter does not prevent the next parameter from being processed.
- Retry count means additional retries after the first attempt. For example, 5 retries means 1 initial attempt + up to 5 retries (6 total attempts).
- Each retry changes thresholding/contrast/inversion/sharpening and OCR page segmentation. This improves resilience but cannot guarantee correct recognition for every font, background, animation, or screen scaling.
- Close the output workbook in Excel while logging; Excel may lock the file and prevent updates.
- This is a screen-reading utility, not a safety instrument. Do not use OCR readings as the sole basis for safety trips or control actions. Verify values against the SCADA source.
