# Import and Export MySQL Data to Excel
Simple Backup, Restore, and Append MySQL Data from Excel using Visual Basic for Applications Code.

## Requirements

- MySQL ODBC Driver:
	- Default Code uses MySQL ODBC 5.1 Driver ([download link](http://dev.mysql.com/get/Downloads/Connector-ODBC/5.1/mysql-connector-odbc-5.1.13-win32.msi)).
	- Other versions can be downloaded from [here](http://dev.mysql.com/downloads/connector/odbc/)

    Note: Code currently supports MySQL ODBC Driver versions 5.1 and 3.51, but other versions will be supported soon. 

## How it works

1. Download and Open "Excel-MySQL-Tools.xlsm" file

2. Inside the '.config' sheet, enter your database connection settings:
	- **Database Host**
	- **Database Name**
	- **Database Username**
	- **Database Password**
	- **Database Connector Driver Name** (See Requirements)

3. Finally, click on one of the action buttons:
	- **Table Structure**: This will create a worksheet for each table in your database, and fill the first row with the table fields. This provides you with empty tables that you can then fill with the data you want to insert into your database.
	- **Backup Data**: This will create a worksheet for each table in your database, and fill each worksheet with the table data, with the first row containing the table field names.
	- **Append Data**: This will take all the data you entered in the worksheets and appends them to your MySQL database.
	- **Restore Data**: This will drop all existing data in your database, then take all the data you entered in the worksheets and appends them to your MySQL database.

## Why use it?

I wanted to create a tool that:

- Allowed me to easily import existing excel data into mysql database for applications I was developing.
- Allow a non-technical business person to easily backup and view a web application data without having to cry for IT who many not be there to help right when they need this data the most :scream:. I recommend creating a database user with read-only privileges to a specific set of tables for this use case.
