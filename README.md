# Automation Tools Using Web Scraping
This project is made to automate common tasks that need to be performed by e-commerce companies.

## Tracking Orders
This script will **scrape through package tracking sites and make requests to their APIs to track 
the status of the packages**. Instead of manually checking to see if each order has been successfully shipped to customers 
or not, the script will automate this task. This will also simplify the process of determining potential shipping errors 
of customer products. 

**NOTE:** The script may have trouble updating the Excel database sheet if it is open by you or any others you have
shared the Excel sheet with when the script finishes running. To make sure the script successfully updates the Excel 
sheet, please make sure no one has the Excel sheet you are updating open while running the script.

## Customer Info
This script will **take existing customer data on online platforms (in this case, Shopify) and populate the correct 
columns in the Excel database**. After exporting the customer data onto a CSV file, it will automate the task of 
extracting relevant customer information and migrating it onto the centralized Excel database sheet. 
