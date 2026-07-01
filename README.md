# E-Commerce Workflow Automator

A lightweight, continuous Python backend service designed to automate manual data entry in high-volume e-commerce product pipelines.

## Overview

In a product pipeline managing thousands of physical items, tracking statuses via manual spreadsheet updates creates a massive administrative bottleneck. This project eliminates that bottleneck by bridging the gap between local file management and the master tracking database.

This script runs as a background polling service. It recursively monitors a designated network directory for new product folders (named by SKU/Barcode). When a photographer or team member drops images into a matched folder, the system automatically patches the `Master.xlsx` database in real-time, updating the product's status and preserving all existing Excel formulas and formatting.

## Technical Implementation

* **Targeted Directory Scanning:** Uses `os.walk` to recursively scan nested directory structures.
* **Network Optimization:** Instead of making thousands of slow network calls across SMB protocols, the script compiles existing directory names into a localized Python set *before* cross-referencing the database, ensuring near-instant execution.
* **Dynamic Column Mapping:** Automatically searches the spreadsheet headers to dynamically locate the target columns (e.g., "Barcode" and "Notes (Sabbir)"), preventing the script from breaking if columns are shifted or added later.
* **Non-Destructive Editing:** Utilizes `openpyxl` rather than `pandas` to ensure cell colors, conditional formatting, and other existing spreadsheet UI elements remain fully intact upon saving.
* **Catch-All Logic:** Intelligently checks if a status is *not* complete, rather than hardcoding a strict list of allowed statuses, making the script highly fault-tolerant to new data entry variations.

## Setup & Installation

**1. Clone the repository:**  
   ```bash
   git clone [https://github.com/YourUsername/ecommerce-workflow-automator.git](https://github.com/YourUsername/ecommerce-workflow-automator.git)
    cd ecommerce-workflow-automator
   ```
   
**2. Install dependencies:**  
   This project requires Python 3.x and openpyxl.  
   ```bash
   pip install openpyxl
   ```
   
**3.  Configure your paths:**  
   Open the main Python script and update the configuration variables to match your local or network environment:  
   ```bash
     EXCEL_FILE_PATH = r"C:\Path\To\Your\Database.xlsx"  
     WATCH_DIRECTORY = r"Z:\Your\Network\Drive"
     POLL_INTERVAL = 300 # Set interval in seconds
```
**4. Run the file:**
  ```bash
    python automator.py
```
