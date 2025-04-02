# Python Update List

This app helps save me hours of work by automating an update task.  
**NOTE:** All of the company information and table names have been changed to maintain confidentiality.

## Background

In my current role, I maintain a **Purchase Requisition** web application. Users fill out a form and select items from a list, which is populated from a **MS SQL Server** table. Every month, I receive an `.xlsx` file from another department that contains the list of items that are currently live in SAP. My responsibility is to update the web app to reflect the new data.

It is important to note that only items with the status **"Open"** are to be displayed on the site. The items change names, descriptions, and status month-to-month, so the process becomes very labor-intensive.

To solve this problem, I created a Python script that automates the following steps:
- It compares the current list from the SQL database with the new data from the Excel file.
- Archives the "closed" items to preserve the old data.
- Updates the web app by uploading the new table with the correct information.

This script significantly reduces the manual work, allowing me to automate the update process each month instead of manually checking and updating the database.

---

## Process Overview

1. **Download the Excel Table**:  
   Every month, I must download the latest item list from SAP in `.xlsx` format from the relevant department. I then rename this file to **newTable.xlsx** for consistency.

2. **Prepare the Script**:  
   The `updateItemList.py` script needs to be in the same directory as the **newTable.xlsx** file. This ensures the script can easily access and process the latest data from the Excel file.

3. **Testing on the Development Server (dev)**:  
   The script first connects to our **development server** ("dev"). This is done for safety and redundancy, as the development server is a backup. Running the script on dev allows me to check that everything works as expected before applying the update to production.

4. **Production Update**:  
   If the script runs successfully on the development server, I can then rerun the script, inputting **"prod"** to update the **production server**. This ensures the changes are applied to the live web application.

---

## Features

- **Automated Comparison**: Compares the old item list with the new one from SAP.
- **Archiving**: Archives the "closed" items for future reference.
- **Updating**: Inserts new items into the database while ensuring that only **"Open"** items are active on the web application.

---

## How It Works

### Step-by-Step Process

1. **Download and Rename the Excel File**:  
   Download the Excel file from the relevant department, and rename it to **newTable.xlsx**. Place this file in the same directory as the `updateItemList.py` script.

2. **Load Config and Excel File**:  
   The script loads a configuration file (`config.json`) containing connection details and the new item list from the Excel file (`newTable.xlsx`).

3. **Database Connection**:  
   It connects to either a **development** or **production** database based on user input:
   - First, the script connects to the **development server** ("dev") to perform a dry run. This is done to test the update process and ensure no issues arise.
   - If successful, the script is rerun with the input **"prod"** to apply the update to the live **production server**.

4. **Item List Comparison**:  
   The script fetches the current item list from the database and compares it with the new list from the Excel file, identifying items that are missing from the new list.

5. **Archiving Closed Items**:  
   The items that no longer appear in the new list are archived by inserting them into an **ItemListArchive** table.

6. **Data Deletion**:  
   The script deletes the old item list from the main table (`ItemList`).

7. **Inserting New Items**:  
   It inserts the updated item list into the **ItemList** table, ensuring that all items are listed as **Open**.

---

## Technologies Used

- **Python**: Used for scripting and automating the process.
- **OpenPyXL**: For reading and writing Excel files.
- **pyODBC**: For interacting with the MS SQL Server database.

---

## Benefits

- **Time-Saving**: What used to take several hours each month is now fully automated.
- **Error Reduction**: By automating the comparison and update process, human error is minimized.
- **Data Integrity**: The script ensures that only the correct items (those with the status "Open") are included in the web application, and that "closed" items are archived for redundancy.

---

## Conclusion

This Python script automates the tedious process of updating the web application’s item list each month, allowing me to save hours of manual work. By leveraging Python and SQL, I was able to streamline the workflow while maintaining data accuracy and consistency. Testing on the development server first ensures that no issues arise before updating the production system, adding an extra layer of safety.
