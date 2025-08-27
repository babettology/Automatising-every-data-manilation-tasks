DPPH Bridge VBA Templates

INDEX 
•	Executive Summary
•	VBA Workflow
•	VBA Templates 
o	A. Table in format 
o	B. Table Overall (DSP+AmFlex) 
o	C. Table DSP 
o	D. Table Flex 
o	E. Table UTR / Delivery Fails
o	F. Table OTR / Dispatch Fails
o	G. Table FTDS 
o	H. Create all tables
o	I. Create DPPH Bridge
o	J. Export into a word document

EXECUTIVE SUMMARY 

What is a VBA
Visual Basic Application is a programming language developed for Excel. 
This document groups useful templates that allow automation, standardisation and streamlining reports to identify potential outliers and issues.
	The templates are made for the metrics in the DPPH_AZ available at : https://perfectmile.amazon.com/dashboards/azanni/region/EU/daily?tab=c9ddad9a-1c06-436e-9c6e-1a479f9f01ab&start-date=2025-08-10&end-date=2025-08-16. 
	Before deploying VBAs, copy the USER INPUT Template as a new sheet. Make sure the sheet name is USER INPUT. The sheet name is case sensitive. 
 
VBA Workflow: 
1.	Download the desired week/Day from DPPH Dashboard on PM
2.	Open the Excel .xls format containing the data
3.	Copy the table into a normal excel finle (.xlsx), make sure no cells are merged.
4.	Click on the "Developer" tab in the upper ribbon 
o	Note: If you don't see the Developer tab, go to File > Options > Customise Ribbon and check the "Developer" box
5.	Click on "Macros"
6.	Type any title in the upper search bar and press Enter
7.	When the VBA editor opens, select all existing code (if any) and delete it
8.	Copy and paste the desired VBA code from this document
9.	Close the editor by clicking the red X in the top-right corner
10.	Return to Excel and click on "Macros" again
11.	Select CreateDPPHBridge 
12.	Click "Run"

What happens next?
After running the macro, Excel will automatically process your data according to the specific template you selected. The results will appear in your workbook in new worksheets. 
 
