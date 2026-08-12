# Pending Order Report Dashboard | DKSH Automation

This project is a React-based interactive web dashboard built to automate the processing, filtering, deduplication, and pivot summarizing of DKSH order reports. It is optimized to run seamlessly in modern web browsers and can be deployed directly to Vercel.

## 🚀 Deployment to Vercel

The root of this project contains a standard React-Vite SPA configuration which is native to Vercel. 
To deploy:
1. Push this folder to a GitHub, GitLab, or Bitbucket repository.
2. Log into your [Vercel Dashboard](https://vercel.com).
3. Import the repository and select **Vite** as the framework template (it is auto-detected).
4. Click **Deploy**. Vercel will build the React bundle on the cloud and host it at a public URL.

---

## 💻 Running Locally (Choose Option A or B)

### Option A: Standalone Sandbox (No Node.js Required - Recommended for offline/immediate use)
We have prepared a standalone dashboard in the `standalone/` directory that runs directly in any browser without needing to install Node.js or `npm`.

1. Run Python's built-in lightweight web server in this directory:
   ```bash
   python -m http.server 8000
   ```
2. Open your web browser and navigate to:
   ```
   http://localhost:8000/standalone/index.html
   ```
   *(Or simply double-click the `standalone/index.html` file in your File Explorer!)*

### Option B: Local Development with Vite (Requires Node.js installed on your PC)
If you decide to install Node.js later:
1. Open your terminal in this project folder.
2. Install the package dependencies:
   ```bash
   npm install
   ```
3. Run the hot-reloading development server:
   ```bash
   npm run dev
   ```
4. Open the development URL (usually `http://localhost:5173`) in your browser.

---

## 📊 Automation & Processing Logic Rules Applied

1. **Exclusion Rules**:
   - Compares the `Payment Status` and `Payment Method` columns.
   - If `Payment Status` is `Pending` (case-insensitive) **AND** the `Payment Method` is **not** `COD` (case-insensitive, e.g. credit card, bank transfer, online payment, etc.), the order is excluded from the final outputs.
   - Rest kept: If status is not pending, or method is COD.
2. **Deduplication Rules**:
   - The same order number can appear multiple times due to multiple line items.
   - Prior to calculating the pivot counts, duplicate rows matching the same `Order Number` (case-sensitive) are removed. Only the first occurrence is kept, meaning each unique order number is counted exactly once.
3. **Pivot Table**:
   - Aggregates the unique, filtered orders grouped by **Nickname / Seller** and **Order Date**.
   - Displays the results in a clean table sorted chronologically and alphabetically.
4. **Yellow Highlighting of Past-Date Orders**:
   - Highlights older date orders in a pastel yellow color for easy visibility.
   - Supports two customizable modes selectable via a toggle:
     - Highlight dates before the latest date found in the upload file.
     - Highlight dates before today's calendar date.
5. **Downloads**:
   - **Export Pivot Table**: Downloads the summarized count table to a clean Excel `.xlsx` file.
   - **Export Cleaned Details**: Downloads the complete list of filtered, deduplicated order rows to a clean Excel `.xlsx` file.
