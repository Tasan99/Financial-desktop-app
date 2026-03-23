# Financial Data Engine

## Overview

Financial Data Engine is a desktop analytics application designed to process, analyze, and visualize financial data from structured Excel files.

The application combines:

* **Python (Flask)** for backend data processing
* **Pandas** for financial computations
* **Electron.js** for desktop application interface

Users can run the application as a standalone executable without requiring Python or any development environment.

---

## Features

* Upload and process Excel-based financial data
* Generate multiple financial reports:

  * Income Statement Comparison
  * Monthly Trended Income Statement
  * Branch Rankings
  * Regional Analysis
  * Operational Metrics Dashboard
* Real-time interaction between frontend and backend
* Fully packaged desktop application (no coding required to run)

---

## Application Architecture

The application follows a **hybrid architecture**:

### Frontend

* Built with HTML, CSS, and JavaScript
* Runs inside an Electron window

### Backend

* Python (Flask API)
* Handles:

  * Data loading
  * Data cleaning
  * Financial calculations
  * Report generation

### Desktop Layer

* Electron manages:

  * Window creation
  * Backend process execution
  * Communication between UI and API

---

## How It Works

1. The user launches the application (`Financial Data Engine.exe`)
2. Electron opens the frontend interface
3. The backend (Python Flask server) is automatically started in the background
4. The user uploads an Excel file
5. The backend processes the data and returns results via API
6. The frontend displays reports dynamically

---

## Installation & Usage

### Option 1 — End User (Recommended)

1. Open the application:

   ```
   Financial Data Engine.exe
   ```
2. Upload your Excel file when prompted
3. Navigate between reports

No additional setup is required.

---

### Option 2 — Development Mode

#### Requirements

* Python 3.x
* Node.js
* npm

#### Install dependencies

```bash
npm install
pip install -r requirements.txt
```

#### Run application

```bash
npm start
```

This will:

* Launch Electron
* Start the Python backend automatically

---

## Project Structure

```
financial-desktop/
│
├── backend/
│   ├── app.py
│   ├── engine.py
│
├── frontend/
│   ├── index.html
│   ├── renderer.js
│   ├── style.css
│
├── main.js
├── package.json
├── build/
├── dist/
```

---

## Backend Endpoints

| Endpoint         | Method | Description         |
| ---------------- | ------ | ------------------- |
| `/api/load_file` | POST   | Loads Excel file    |
| `/api/entities`  | GET    | Returns entity list |
| `/api/report1`   | GET    | IS Comparison       |
| `/api/report2`   | GET    | Monthly IS          |
| `/api/report3`   | GET    | Rankings            |
| `/api/report4`   | GET    | Regional Analysis   |
| `/api/report5`   | GET    | Metrics Dashboard   |

---

## Known Limitations

* The application expects a specific Excel format:

  * Sheet name: `Raw Data`
  * Structured financial layout
* Large files may increase processing time
* Backend runs locally (not cloud-based)

---

## Future Improvements

* Export reports (PDF / Excel)
* Cloud deployment option
* Authentication system
* Performance optimization for large datasets

---

---

## Notes

This application demonstrates how financial analytics pipelines can be integrated into a fully functional desktop environment using modern tools.

---
## Important Note – Closing the Application Properly

In some cases, the application backend may continue running in the background after closing the main window.

If you experience issues such as:

* The app not reopening
* Port already in use errors
* System slowing down

Please follow these steps:

1. Open **Task Manager** (Ctrl + Shift + Esc)
2. Look for:

   * `python.exe` (development mode)
   * or the application process (packaged version)
3. Select it and click **End Task**

This ensures the backend process is fully terminated.

---

**Tip:** Always try to close the application normally before using Task Manager.
