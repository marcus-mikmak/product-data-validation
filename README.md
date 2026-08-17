# Brand Validation Tools - Excel Add-in

This repository contains the Microsoft Excel Web Add-in for Brand Data Validation. It replaces the legacy Google Apps Script (GAS) tool, providing a native, centralized solution within Microsoft 365 for importing, mapping, validating, and exporting product data.

## 🏗️ Architecture

This tool is built as an **Office Web Add-in**. 
* **Frontend/Backend:** Standard HTML, CSS, and JavaScript (utilizing the `Office.js` and Excel JavaScript APIs).
* **Hosting:** The codebase is hosted statically via **GitHub Pages**.
* **Deployment:** Excel loads the tool via a centralized XML Manifest (`manifest.xml`), meaning updates pushed to the `main` branch are instantly reflected for all users without requiring re-installation.

## 📂 Repository Structure

```text
my-excel-addin/
├── manifest.xml                 # The configuration file for M365 deployment
├── README.md                    # Project documentation
└── src/
    ├── js/
    │   └── excelAddin.js        # Core Excel API logic and batch processing
    ├── html/
    │   ├── Commands.html        # Invisible file for Ribbon button routing
    │   ├── ImportDialog.html    # Upload & Mapping modal for brand files
    │   ├── MinotaurImportDialog.html  # Minotaur existing data upload modal
    │   ├── MinotaurExportDialog.html  # Minotaur export mapping modal
    │   └── ProgressSidebar.html # Taskpane for manual image validation
    └── assets/                  # (Add your 64x64 and 128x128 icon PNGs here)
```

## ✨ Features

* **Brand File Import:** Ingests CSV/XLSX files locally, maps external brand columns to internal template headers, and bulk-writes the data into Excel.
* **Image Validation:** Checks image URLs (via partial-byte HTTP requests) directly from the browser to ensure product images render correctly before export.
* **Conditional Exports:** Filters the grid based on the `INTERNAL_Data_Readiness_Score` to generate isolated "Good Data" or "Review Data" tabs.
* **Minotaur Integration:** Maps template data to Minotaur-specific headers and exports clean sheets for system ingestion.

---

## 🚀 Deployment Guide

### 1. Activate GitHub Pages
To make the add-in accessible to Excel, the code must be served via HTTPS.
1. Go to repository **Settings** > **Pages**.
2. Set the Source to **Deploy from a branch**.
3. Select the `main` branch and `/root` folder, then click **Save**.
4. Copy the generated live URL (e.g., `https://your-org.github.io/repo-name/`).

### 2. Configure the Manifest
Update the `manifest.xml` file to point to your live GitHub Pages environment. 
* Do a Find & Replace for `https://YOUR_GITHUB_USERNAME.github.io/YOUR_REPO_NAME/` and replace it with your actual GitHub Pages URL.

### 3. Local Testing (Sideloading)
For implementation testing before an org-wide release:
1. Open Excel on the Web (Office.com).
2. Go to **Home** > **Add-ins** > **More Add-ins**.
3. Click **Manage your apps** (bottom left).
4. Select **Upload custom app** -> **Upload from my device**.
5. Select the local `manifest.xml` file.
*Note: If sideloading is disabled, contact IT to request developer permissions or a scoped deployment.*

### 4. Organization Deployment (Production)
To deploy this tool to the entire implementation team:
1. Go to the **Microsoft 365 Admin Center**.
2. Navigate to **Settings** > **Integrated Apps**.
3. Click **Upload custom apps**.
4. Upload the `manifest.xml` file and assign it to the relevant user groups.

---

## 🛠️ Development Notes

* **Batch Processing:** Excel JS API requests are batched using `context.sync()`. For large files, logic is chunked to prevent browser freezing.
* **Cross-Origin Resource Sharing (CORS):** The Image Validation tool relies on client-side `fetch()` requests. Some external CDNs may block these requests via CORS. 
* **State Management:** Because this runs locally in the user's browser engine (Edge WebView2), data chunks are stored in memory during processing rather than cached on an external server.
