# ULJK Bloomberg Ranking System

Full-stack Bloomberg-style stock ranking dashboard for comparing two daily Excel files:

1. Yesterday file
2. Today file

The app computes Top 10 and Bottom 10 rank movement per `Index_Clean`, using `Index_Clean + Company Name` as the comparison key.

## Stack

- Frontend: React + Vite
- Backend: FastAPI
- Excel processing: pandas + openpyxl
- API client: axios

## Project Structure

```text
stock-rank-dashboard/
|-- backend/
|   |-- main.py
|   `-- requirements.txt
|-- Daily Data Files/
|   `-- sample Excel inputs
|-- frontend/
|   |-- package.json
|   |-- index.html
|   |-- vite.config.js
|   `-- src/
|       |-- main.jsx
|       |-- App.jsx
|       |-- api.js
|       |-- styles.css
|       `-- components/
|           |-- UploadPanel.jsx
|           |-- KPICards.jsx
|           |-- IndexSelector.jsx
|           `-- RankingTable.jsx
`-- README.md
```

## Backend Setup

Use the existing virtual environment or create one.

Install backend dependencies:

```powershell
.\.venv\Scripts\python.exe -m pip install -r backend\requirements.txt
```

Run the API:

```powershell
.\.venv\Scripts\python.exe -m uvicorn backend.main:app --reload --host 127.0.0.1 --port 8000
```

Backend endpoint:

- `POST /compare`

Form fields:

- `yesterday_file`
- `today_file`

## Frontend Setup

Install frontend dependencies:

```powershell
cd frontend
npm.cmd install
```

Run the React app:

```powershell
npm.cmd run dev
```

Default frontend URL:

- `http://127.0.0.1:5173`

The frontend calls the backend at:

- `http://127.0.0.1:8000`

You can override it with:

```powershell
$env:VITE_API_BASE_URL="http://127.0.0.1:8000"
```

## How To Use

1. Open the frontend in the browser.
2. Upload:
   - yesterday Excel file
   - today Excel file
3. Click `Compare Files`.
4. Select an index from the dropdown.
5. Review:
   - Top 10 table
   - Bottom 10 table
   - KPI cards for the selected index

## Expected Excel Columns

The backend looks for these logical fields:

- `Company Name`
- `Index_Clean`
- `CMP`
- `FS`
- `TS`
- `Total`
- `RSI`
- `30WMA` or `30 WMA`
- `ROCE`
- `ROE`
- `EPS Growth`
- `D/E`

It is tolerant of:

- extra spaces
- grouped Excel headers
- `FY25/FY26` style columns for `ROCE` and `ROE`
- extra unused columns
- sheet name differences

The backend automatically detects the header row by finding a row containing:

- `Company Name`
- `Index_Clean`

## Ranking Logic

All ranking is done per `Index_Clean`.

Comparison key:

- `Index_Clean + Company Name`

This means the same stock in multiple indices is treated separately.

### Top Rank

- Sort by `Total` descending
- Highest score gets `Top Rank 1`
- Next highest gets `Top Rank 2`

### Bottom Rank

- Sort by `Total` ascending
- Lowest score gets `Bottom Rank 1`
- Next lowest gets `Bottom Rank 2`

### Top Change

- If missing yesterday: `New Entry`
- If today rank is smaller than yesterday rank: `Up by X`
- If today rank is larger than yesterday rank: `Down by X`
- If same: `Unchanged`

### Bottom Change

- If missing yesterday: `New Entry`
- If today bottom rank is smaller than yesterday bottom rank: `Down by X`
- If today bottom rank is larger than yesterday bottom rank: `Up by X`
- If same: `Unchanged`

## API Response Shape

The backend returns:

```json
{
  "indices": ["Bank Nifty", "Nifty 50"],
  "data": {
    "Bank Nifty": {
      "summary": {
        "stocks_in_index": 12,
        "new_top10_entries": 2,
        "new_bottom10_entries": 1,
        "top10_unchanged": 3,
        "bottom10_unchanged": 2
      },
      "top10": [],
      "bottom10": []
    }
  }
}
```

## Current Verification

Verified locally:

- backend imports successfully
- backend ranking logic runs against the uploaded sample files
- frontend dependencies install successfully
- frontend production build completes successfully
