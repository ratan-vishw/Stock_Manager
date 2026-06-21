# Stock/Inventory & Bill Manager

A powerful Google Apps Script-based free web application for managing inventory, tracking transactions, and automatically extracting data from bills/invoices using Google's Gemini Vision AI.

## Features

- **Inventory Tracking:** Manage stock going 'In', 'Out', or for 'Repair' across different sites.
- **Dynamic Categories:** Items are organized into categories, each mapping to a dedicated Google Sheet tab (prefixed with `MT.`).
- **AI-Powered Bill Extraction:** Upload images of bills or invoices. The app uses the Gemini 2.5 Flash API to automatically extract line items (Item Name, Quantity, Rate, Unit) and populate the form.
- **Drive Integration:** Uploaded item images and bill receipts are automatically saved to dedicated Google Drive folders (`Inventory_Bills_Uploads` and `Bill Manager Images`), and their links are stored in the database.
- **Transaction Logs:** Keeps a detailed history of every transaction in the `All_Transactions` sheet.
- **Auto-Billing:** Bills can be directly appended to the exact 13-column `Bills` sheet format, complete with PI Numbers, extraction notes, and generated IDs.

## Database Structure

The script expects and manages the following sheets in your Google Spreadsheet:
- **`All_Transactions`**: Logs all inventory movements.
- **`Bills`**: Logs processed bills and extracted invoice items.
- **`SiteNameList`**: A dynamic list of work sites/locations.
- **`MT.[CategoryName]`**: Individual category sheets (e.g., `MT.Hardware`, `MT.Glass`). These use a strict 5-column schema:
  1. Item Name / Description
  2. Current Stock
  3. Unit
  4. Rate (₹)
  5. Image Link

## Setup & Installation

### 1. Prepare your Google Sheet
1. Create a new Google Spreadsheet.
2. (Optional) Create the sheets manually: `All_Transactions`, `Bills`, and `SiteNameList`. The script will automatically generate these and category sheets when needed, but creating them beforehand is fine.

### 2. Add the Apps Script Code
1. In your Google Sheet, click on **Extensions > Apps Script**.
2. Rename the default `Code.gs` (if needed) and paste the contents of the local `Code.js` into it.
3. Create a new HTML file by clicking the **+** icon > **HTML**. Name it **exactly** `Index` (so it creates `Index.html`).
4. Paste the contents of the local `Index.html` into this new file.

### 3. Configure the Gemini API Key
This app requires a Gemini API key to process images of bills.
1. Get an API key from [Google AI Studio](https://aistudio.google.com/).
2. In the Apps Script Editor, go to **Project Settings** (the gear icon on the left).
3. Scroll down to **Script Properties** and click **Add script property**.
4. Set the **Property** to `GEMINI_API_KEY`.
5. Set the **Value** to your actual API key.
6. Click **Save script properties**.

### 4. Deploy as a Web App
1. In the Apps Script Editor, click **Deploy > New deployment** in the top right.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Under **Execute as**, select **Me**.
4. Under **Who has access**, select **Anyone** (or depending on your organization, "Anyone within [Your Org]").
5. Click **Deploy**.
6. Authorize the necessary permissions (Google Sheets, Google Drive, External Requests).
7. Copy the generated Web App URL and share it with your team.

## Usage Guide
- **Add New Items:** If an item doesn't exist, the UI allows you to add it on the fly while doing an "In" transaction. You can even attach a photo.
- **Scan a Bill:** In the form, click on the bill upload section to select a receipt image. The Gemini AI will read it and pre-fill the form with items.
- **Sites:** Start typing a site name in the "Site Location" field. It will auto-suggest based on previously entered sites in the `SiteNameList` sheet.
