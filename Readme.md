# Tracked Uninvoiced Parts Scanner

A desktop application that scans your submitted Parts Unlimited orders and identifies parts that have been **shipped with a tracking number but not yet invoiced**. Results are exported directly to an Excel workbook for easy review.

---

## Download & Install (Windows)

> **No coding or technical knowledge required.** Just download, extract, and run.

### Step 1 — Download the ZIP

Click the file `Find_Uninvoiced_Parts_v2.4.zip` in the file list above and click **Download**.

![Download ZIP](https://raw.githubusercontent.com/rmtodd2/Missing-Invoice-Scanner/master/images/download1.jpg)

### Step 2 — Extract the ZIP

You **must** extract the ZIP before running the app. Running it directly from inside the ZIP will not work.

1. Find the downloaded file on your computer.
2. Right-click the ZIP file and choose **Extract All…**
3. Pick a location (such as your Desktop) and click **Extract**.

> **Tip:** If you double-click the ZIP and see the files inside, you are *browsing* the archive — not extracting it. The app will not run correctly this way. Always right-click → **Extract All…** first.

### Step 3 — Run the App

Inside the extracted folder, double-click **`Find_Uninvoiced_Parts.exe`** to launch the app.

**Windows SmartScreen warning?** If Windows shows a blue "Windows protected your PC" popup, click **More info**, then click **Run anyway**. This is normal for unsigned apps.

---

## Download & Install (macOS)

### Step 1 — Download the ZIP

Click the file `Find_Uninvoiced_Parts_v2.4(Mac).zip` in the file list above and click **Download**.

![Download ZIP](https://raw.githubusercontent.com/rmtodd2/Missing-Invoice-Scanner/master/images/download1.jpg)

### Step 2 — Extract the ZIP

1. Locate the downloaded `.zip` file
2. Double-click it to extract. A new folder will appear next to it.

> **Tip:** If you see files inside a window but there is no new folder on disk, macOS may not have extracted the archive. Try right-clicking the ZIP and choosing **Open With → Archive Utility** to ensure it fully extracts.

### Step 3 — Run the App

1. Open the extracted folder
2. Double-click the **`.app`** file

---

### Mac Security Warning

macOS may show a message like:

> "This app can't be opened because it is from an unidentified developer."

To open the app:

1. Right-click (or Control-click) the `.app` file
2. Click **Open**
3. Click **Open** again in the popup

You only need to do this the first time.

---

### Important (Mac Only)

Mac users **must choose a save location** for the output file.

Click **Browse…** and select a folder (such as Desktop or Documents).

If you do not do this, you may get a **permission error** when the app tries to save the Excel file.

---

## Using the App

### 1. Enter Your Credentials

| Field | What to enter |
|---|---|
| **Dealer #** | Your dealer code from Parts Unlimited |
| **User ID** | Your Parts Unlimited login username |
| **Password** | Your Parts Unlimited login password |

### 2. Set Your Date Range

- **Oldest Date to Search** — The earliest order date you want to check.
- **Newest Date to Search** — The most recent order date to include.

By default, the date range is set to the **previous full month**. For example, if you open the app in April, it will search all of March. You can adjust the dates to any range you need.

The app will only scan orders submitted **within this date range**.

### 3. Choose Where to Save the Output

The **Output File** field shows where your Excel report will be saved.

- On **Windows**, the default location works fine
- On **Mac**, you should always click **Browse…** and select a location manually

### 4. Start the Scan

Click **Start Scan**. The console at the bottom of the window will show live progress — which pages are being scanned, how many orders were checked, and what was found.

When finished, a popup will confirm success and show you where the Excel file was saved.

---

## Understanding the Output

The generated Excel file has **two sheets**:

### Sheet 1 — Summary

A simple list of every order that has at least one tracked-but-uninvoiced part:

| Column | Description |
|---|---|
| Dealer # | Your dealer code |
| Order Number | The Parts Unlimited order number |

Duplicate orders are automatically removed.

### Sheet 2 — Full Output

Every tracked, uninvoiced line item with full detail:

| Column | Description |
|---|---|
| Dealer # | Your dealer code |
| Order Number | Order number |
| Order Status | Current status of the order |
| Submitted Date | Date the order was placed |
| Part | Part number (formatted) |
| Description | Part description |
| Brand | Part brand |
| Line Status | Status of the specific line item |
| Quantity | Qty originally ordered |
| Shipped | Qty that has shipped |
| Invoiced | Qty that has been invoiced |
| Tracking | Tracking number(s) |
| Ship Date | Date the item shipped |
| Ship Via | Shipping carrier |
| Location | Warehouse location code |

---

## Advanced Settings

> **Most users should leave these at their default values.** These control how the app talks to the Parts Unlimited server.

| Setting | Default | What it does |
|---|---|---|
| Request Delay (sec) | `0.35` | Pause between each request to avoid overloading the server |
| Retry Backoff (sec) | `2.0` | How long to wait before retrying a failed request |
| Rate Limit Max Requests | `900` | Max requests allowed within the time window |
| Rate Limit Window (sec) | `300` | The rolling time window (in seconds) for rate limiting |

---

## Troubleshooting

**Login failed: invalid credentials**
Double-check your Dealer #, User ID, and Password. These are the same credentials you use to log in to Parts Unlimited.

**The scan runs but finds nothing**
All shipped orders in your date range have been fully invoiced — that's a good thing! Try widening your date range if you expected to find something.

**The app won't open / crashes immediately**
Make sure you extracted the ZIP before running the app. On Windows, right-click the ZIP → **Extract All…** → then open the extracted folder and run the `.exe` from there.

**Windows says the file is dangerous**
This is a SmartScreen warning for unsigned software. Click **More info → Run anyway** to proceed.

**Mac says the app cannot be opened**
Right-click the `.app` file → **Open** → click **Open** again.

**Mac permission error when saving file**
Make sure you selected a save location using the **Browse…** button before starting the scan.
