# MavsAI_IT_Inventory_Scripts
These are the scripts I used to collect device configuration details from employees using various operating systems. Then, I converted them into a nicely-formatted XLSX file.

The "Data_Collection_*" files are the main scripts that I ran on every employee's devices. This created a TXT file on their desktop. I collected all such TXT files from everybody.

Then I ran the "Data_to_XLSX_Conversion_Script.ps1". This targets a specific folder where all the TXT files were saved. It takes all of them, creates a basic CSV out of them. Then takes that CSV and makes it into a well-formatted XLSX file.
