QuickBooks-to-Business-Central

This repo is mainly a set of QuickBooks Advance Reporting reports to export data out of QuickBooks Desktop and import into into a few tables in Business Central.
The QVW files can be copied to your QuickBooks data directory where there should be a folder named "Company Name" + Advanced Reporting. save the QVW files into that folder. 

You can run and modify the reports inside QuickBooks Advanced Reporting and export your data via the export to excel button.
I had to save the files has XLSX for BC to import them, QuickBooks will save them as XLS or CSV.

The .AL file is a set of procedure that you can call from any pages.
The Customer, Vendor, Contacts, Ship-Tos and Items do not require new tables to be created but you may need to add fields or comment out some of the code if you choose not to use them.

The Estimate, Invoice and Purchase order are tables and pages we created to display the imported content. They are not included in this repo at this time.

I hope this may come in handy for someone as we spent a lot of time messaging the QB data to get it in to BC.
Cheers
