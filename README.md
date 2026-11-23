# Screen Data Logger

**Automated SCADA Screen Monitoring and Excel Data Logging Using OCR**

Screen Data Logger is a Windows desktop application developed in Python to automatically capture numerical readings displayed on SCADA systems, industrial monitoring software, and other computer screens.

It uses Optical Character Recognition (OCR) to extract values from user-selected screen regions and records them in an Excel workbook along with timestamps.

The application is designed to reduce manual data collection, simplify industrial process monitoring, and maintain historical operating data.

---

## Key Features

- **Multiple Parameter Monitoring:** Configure multiple parameters from different locations on the screen.
- **Custom Screen Selection:** Select an individual screen region for each parameter.
- **Automatic OCR Recognition:** Extract numerical values from digital screen displays.
- **Sequential Processing:** Process configured parameters one by one during each logging cycle.
- **Configurable Logging Interval:** Set the time interval between successive logging cycles.
- **Automatic OCR Retry:** Retry unreadable or failed OCR readings using different image-processing techniques.
- **Configurable Retry Count:** Set the number of additional OCR retries. Default: 5 retries after the initial attempt.
- **Image Preprocessing:** Apply image resizing, grayscale conversion, contrast enhancement, sharpening, thresholding, and inversion.
- **Automatic Excel Logging:** Store readings in `.xlsx` format using OpenPyXL.
- **Timestamp Recording:** Record each logging cycle with a timestamp.
- **OCR Failure Tracking:** Record failed readings as blank values with an OCR failure status.
- **Configuration Persistence:** Retain parameter settings and application configuration between sessions.

---

## Application Workflow

1. Add the required monitoring parameters.
2. Select a screen region for each parameter.
3. Configure the logging interval.
4. Set the OCR retry count.
5. Test OCR recognition for each parameter.
6. Start the logging process.
7. The application sequentially captures each parameter, extracts its value, and saves the readings to Excel.

### OCR Retry Logic

When an OCR attempt returns an unreadable value or an error, the application automatically retries using alternative image preprocessing and recognition configurations.

For example, with the retry setting configured to 5:

- Initial OCR attempt: 1
- Additional retries: Up to 5
- Maximum total attempts: 6 per parameter

If all attempts fail, the value is left blank and the corresponding OCR status is recorded as failed. Previous readings are never automatically reused.

---

## Application Screenshots

Screenshots can be added here to demonstrate:

- Main application interface
- Parameter configuration
- Screen-region selection
- OCR testing
- Excel logging output

---

## Technology Stack

| Component | Technology |
|---|---|
| Programming Language | Python |
| Desktop GUI | PySide6 |
| OCR Engine | Tesseract OCR |
| Image Processing | Pillow |
| Excel Integration | OpenPyXL |
| Platform | Windows |

---

## Installation

### Prerequisites

Install the following software before running the application:

1. Python 3.10 or newer
2. Tesseract OCR for Windows

Tesseract OCR is available here:

https://github.com/UB-Mannheim/tesseract/wiki

The application looks for Tesseract at the standard installation location:

`C:\Program Files\Tesseract-OCR\tesseract.exe`

If Tesseract is installed elsewhere, configure the `TESSERACT_CMD` environment variable with the full executable path.

### Setup Instructions

**Step 1: Download the project**

Clone the repository:

```bash
git clone https://github.com/divyeshghoghari/screen-data-logger.git
```

Alternatively, download the repository as a ZIP file and extract it.

**Step 2: Open the project directory**

```bash
cd screen-data-logger
```

**Step 3: Install dependencies**

```bash
pip install -r requirements.txt
```

**Step 4: Run the application**

```bash
python screen_data_logger.py
```

Alternatively, Windows users can run:

`Run_Screen_Data_Logger.bat`

---

## Usage Instructions

### 1. Add Parameters

Add the required process parameters, such as:

- Gas pressure
- Gas flow
- Temperature
- Compressor discharge pressure
- Vacuum pressure
- Methane concentration
- CO2 concentration

### 2. Select Screen Regions

Select each parameter individually and draw a rectangle around the numerical value displayed on the screen.

### 3. Configure Logging

Set the desired logging interval and OCR retry count.

Example configuration:

| Setting | Example |
|---|---|
| Number of parameters | 10 |
| Logging interval | 30 seconds |
| OCR retries | 5 |
| Processing mode | Sequential |
| Output format | Excel (.xlsx) |

### 4. Test OCR

Use the OCR testing function to verify that each selected region produces the expected numerical value.

### 5. Start Logging

Start the monitoring process. The application will capture all enabled parameters and save the results to the selected Excel workbook.

---

## Excel Output

The application generates an Excel workbook containing:

- Timestamp
- Individual parameter readings
- Individual parameter OCR statuses

Each logging cycle is saved as a new row.

Example:

| Timestamp | Gas Pressure | Temperature | Flow Rate |
|---|---:|---:|---:|
| 10:00:00 | 0.52 | 35.2 | 420 |
| 10:00:30 | 0.53 | 35.4 | 422 |
| 10:01:00 | 0.51 | 35.3 | 418 |

*Illustrative values only.*

---

## Building a Windows Executable

The project includes a batch file for creating a standalone Windows executable using PyInstaller.

Run:

`Build_Windows_EXE.bat`

The generated executable will be available inside the `dist` directory.

Tesseract OCR may still need to be installed separately unless it is packaged and configured with the executable.

---

## Project Structure

```text
screen-data-logger/
│
├── screen_data
