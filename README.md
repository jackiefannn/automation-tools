# Automation Tools Using Web Scraping
This project is made to automate common tasks that need to be performed in companies (mainly e-commerce businesses).

## Tracking Orders
Instead of manually checking to see whether or not each order has been shipped to customers 
or not, the script will automate this task. It will **scrape through package 
tracking sites and make requests to their APIs to track the status of the packages**.  


**NOTE:** The script may have trouble updating the Excel database sheet if it is open by you or any others you have
shared the Excel sheet with while the script is running. To make sure the script successfully updates the Excel sheet, 
please make sure no one has the Excel sheet you are updating open while running the script.

## Customer Info
This script will automate the task of storing customer information on websites into an Excel database. It will 
**take existing customer data on online platforms (in this case, Shopify) and populate the correct columns in the Excel 
database**.