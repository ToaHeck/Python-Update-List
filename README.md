# Python Update List
A Python script to automate monthly updates to a web application's item list, saving hours of manual work while ensuring data accuracy.

**Note**: Company details and table names are anonymized for confidentiality.

## Background
In my role maintaining a Purchase Requisition web app, I update an MS SQL Server table that populates an item list for user forms. Each month, I receive an Excel file from another department with the latest SAP-derived item list. Manually updating this—tracking name/description changes and "Open"/"Closed" statuses—was time-consuming. This script automates the process by:
- Comparing the current database list with the new Excel data.
- Archiving "Closed" items for historical reference.
- Refreshing the live table with "Open" items.

## Features
- **Automated Comparison**: Identifies differences between old and new item lists.
- **Archiving**: Preserves outdated items in an archive table.
- **Database Updates**: Deletes old data and inserts new "Open" items.
- **Dev/Prod Support**: Tests updates on a development server before applying to production.

## How It Works
1. **Download Excel**: Save the latest SAP item list as `newTable.xlsx` in the script directory.
2. **Load Data**: Reads `config.json` for database credentials and `newTable.xlsx` for item data.
3. **Connect to DB**: Prompts user for "dev" (testing) or "prod" (live) server connection.
4. **Compare Lists**: Uses Python sets for efficient lookup of items to archive or update.
5. **Archive Closed Items**: Inserts missing items into `ItemListArchive` with "Closed" status.
6. **Update Live Table**: Clears `ItemList` and inserts new items with "Open" status.

## Technologies
- **Python**: Core scripting language.
- **OpenPyXL**: Excel file parsing.
- **pyODBC**: MS SQL Server interaction.


## Flow Diagram
![Flow diagram of how the application works](https://github.com/ToaHeck/Python-Update-List/blob/main/img/updateItemList-Diagram.png)
