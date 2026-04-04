# Tracked Uninvoiced Parts Scanner

This application scans submitted orders and identifies parts that have been **shipped with tracking but not yet invoiced**. Results are exported to an Excel file for easy review.

---

## Getting Started

1. Launch the application  
2. Enter your login credentials  
3. Select your date range  
4. Choose where to save the output file  
5. Click **Start Scan**

---

## Input Fields

### Dealer #
Enter your dealer code provided by Parts Unlimited.

---

### User ID
Your login username.

---

### Password
Your login password.

---

### Oldest Date to Search
The earliest date you want to include in the scan.

---

### Newest Date to Search
The most recent date you want to include in the scan.

The app will only scan orders **within this date range**.

---

### Output File
Choose where the Excel file will be saved.

- Click **Browse** to select a location  
- The file will be created automatically after the scan completes  

---

## Output

The generated Excel file contains two sheets:

### Summary
- Dealer #
- Order Number  
- Duplicate orders removed  

---

### Full Output
Detailed results including:
- Order information  
- Part numbers  
- Tracking numbers  
- Quantities shipped vs invoiced  

---

## Console Output

The console window shows:
- Scan progress  
- Pages being processed  
- Final results summary  

---

## Advanced Settings (Leave as Default)

These settings control how the app communicates with the server.  
**Most users should NOT change these.**

---

### Request Delay (sec)
- Default: `0.35`
- Adds a delay between requests to avoid overloading the server  

---

### Retry Backoff (sec)
- Default: `2.0`
- Time to wait before retrying failed requests  

---

### Rate Limit Max Requests
- Default: `900`
- Maximum number of requests allowed within the time window  

---

### Rate Limit Window (sec)
- Default: `300`
- Time window used for rate limiting  

---
