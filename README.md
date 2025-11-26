# Automated-Certificate-Distributor

A Node.js automation tool designed to streamline the mass distribution of event certificates. This script parses recipient data from Excel, manages file organization, and utilizes the **SMTP protocol** (via `nodemailer`) to securely dispatch personalized attachments to users.

---

## 🛠️ Tech Stack
* **Runtime:** Node.js
* **Email Protocol:** SMTP (via `nodemailer`)
* **Data Parsing:** Excel (XLSX) handling
* **File System:** `fs` module for file renaming and verification

## 📂 Prerequisite Folder Structure
For the script to function correctly, your directory must be organized exactly as follows:

```text
├── 📂 certificates/         <-- Folder containing raw certificate images/PDFs (from Canva/Photoshop)
├── 📄 recipients.xlsx       <-- Excel sheet with user data
├── 📄 index.js              <-- Main automation script
└── 📄 package.json

## 📋 Input Requirements

### 1. Recipient Data (`recipients.xlsx`)
The automation relies on a clean Excel sheet located in the root directory.
* **Format:** `.xlsx`
* **Structure:** It must contain **two columns only**:
    * **Column A:** Name of the Recipient
    * **Column B:** Email Address
* **Constraint:** Ensure there are **no empty cells** in these columns to prevent execution errors.

### 2. Certificate Files
* **Source:** Certificates generated from tools like Canva or Photoshop.
* **Location:** All certificate files (PDF or Image) must be placed inside the folder named `certificates`.
* **Note:** The script will process files in this folder alphabetically or based on the directory order, so ensure your Excel list matches the file order if names are not yet assigned.

---

## 🚀 Usage Procedure

This script is designed with a **two-step safety workflow**. You will need to modify `index.js` slightly (by commenting/uncommenting blocks) to switch between these modes.

### Phase 1: File Renaming
*Objective: Assign recipient names to the raw certificate files.*

1.  Open `index.js`.
2.  **Enable:** The "Renaming Section" code block.
3.  **Disable (Comment out):** The "Sending Section" (Nodemailer logic).
4.  Run the script:
    ```bash
    node index.js
    ```
5.  **Verification:** Go to the `certificates` folder and confirm that all files have been renamed to the correct user names.

### Phase 2: Email Distribution
*Objective: Send the verified certificates to users via SMTP.*

1.  Once you have verified the filenames are correct, open `index.js`.
2.  **Disable (Comment out):** The "Renaming Section".
3.  **Enable:** The "Sending Section".
4.  Run the script to start sending emails:
    ```bash
    node index.js
    ```

---

## ⚠️ Important Notes
* **SMTP Config:** This project uses `nodemailer`. Ensure your SMTP credentials are valid.
* **Data Match:** The number of files in the `certificates` folder **must match** the number of rows in `recipients.xlsx`.
