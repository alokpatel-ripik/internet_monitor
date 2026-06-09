# Internet Connectivity Monitor

A production-ready Python application that monitors internet connectivity 24/7, logs results to Excel, and uploads daily reports to AWS S3 with full audit tracking.

## Features

- **Continuous monitoring** — checks internet connectivity every **15 minutes** on clock boundaries (`:00`, `:15`, `:30`, `:45`)
- **Excel logging** — one worksheet per day (`YYYY-MM-DD`) with `Timestamp` and `Connected` (True/False)
- **7-day retention** — automatically removes worksheets older than 7 days
- **Scheduled S3 upload** — exports and uploads the previous day's report at **00:30 AM**
- **Failed upload queue** — stores failed reports in `pending_uploads/` and retries every 15 minutes while online
- **Upload audit trail** — permanent `Upload_Audit` sheet tracks every upload attempt (SUCCESS / PENDING)
- **Resilient** — exception handling ensures the monitor keeps running even if checks or uploads fail

## How It Works

```
Every 15 min  →  Check google.com  →  Save to today's Excel sheet
00:30 AM      →  Export yesterday's sheet  →  Upload to S3
Upload fails  →  Save to pending_uploads/  →  Retry when internet returns
```

### Connectivity check

Sends an HTTP GET request to `https://www.google.com` with a 10-second timeout. Returns `True` if the response is successful (2xx), `False` otherwise.

### Daily upload (00:30 AM)

At the first check on or after **00:30**, the previous day's worksheet is exported as `YYYY-MM-DD.xlsx` and uploaded to S3. If the upload fails, the file is moved to `pending_uploads/` and a `PENDING` row is added to `Upload_Audit`.

### Pending upload retry

While internet is connected, every 15-minute cycle scans `pending_uploads/` and retries failed uploads. On success, the file is deleted and a new `SUCCESS` audit row is appended.

## Excel Workbook Structure

**File:** `Internet_Monitor.xlsx` (created automatically in the project folder)

### Daily sheets (`YYYY-MM-DD`)

| Timestamp           | Connected |
|---------------------|-----------|
| 2026-06-09 14:00:00 | TRUE      |
| 2026-06-09 14:15:00 | FALSE     |

### Upload_Audit (permanent)

| Report Date | Upload Attempt Time | Status  |
|-------------|---------------------|---------|
| 2026-06-08  | 2026-06-09 00:30:05 | PENDING |
| 2026-06-08  | 2026-06-09 08:15:00 | SUCCESS |

Audit rows are **append-only** — records are never overwritten.

## Requirements

- Python 3.9+
- Dependencies: `openpyxl`, `requests`, `boto3`

## Installation

```bash
git clone https://github.com/alokpatel-ripik/internet_monitor.git
cd internet_monitor

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Configuration

Set AWS credentials and S3 settings via environment variables:

```bash
export S3_BUCKET=your-bucket-name
export AWS_ACCESS_KEY_ID=your-access-key
export AWS_SECRET_ACCESS_KEY=your-secret-key
export AWS_REGION=us-east-1

# Optional — default: internet-monitor/
export S3_PREFIX=internet-monitor/
```

> **Never commit credentials.** Use environment variables or a `.env` file (already in `.gitignore`).

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `S3_BUCKET` | Yes | — | Target S3 bucket |
| `AWS_ACCESS_KEY_ID` | Yes | — | AWS access key |
| `AWS_SECRET_ACCESS_KEY` | Yes | — | AWS secret key |
| `AWS_REGION` | No | `us-east-1` | AWS region |
| `S3_PREFIX` | No | `internet-monitor/` | S3 key prefix |

If `S3_BUCKET` is not set, uploads fail gracefully and reports are queued in `pending_uploads/`.

## Usage

```bash
source .venv/bin/activate
python3 internet_monitor.py
```

Stop with `Ctrl+C`.



## Project Files

| File / Folder | Description |
|---------------|-------------|
| `internet_monitor.py` | Main application |
| `requirements.txt` | Python dependencies |
| `Internet_Monitor.xlsx` | Connectivity data + audit log (auto-created) |
| `internet_monitor.log` | Runtime log file |
| `pending_uploads/` | Queued reports awaiting S3 upload |

## S3 Layout

```
s3://your-bucket/internet-monitor/2026-06-08.xlsx
s3://your-bucket/internet-monitor/2026-06-09.xlsx
```

## Known Limitations

- **Laptop sleep:** If the machine sleeps, missed checks are backfilled on wake using the **current** internet status at wake time — not the actual status during sleep.
- **True 24/7 monitoring** requires the machine to stay awake or the app to run as a background service (e.g. launchd).

## License

MIT
