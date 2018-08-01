# QuickBooks-To-CSV Summary
A C# console application created using the QuickBooks SDK; queries the current company file and writes results to csv

# Value of Project
While QuickBooks is excellent for creating and storing financial records, very little affordable support exists for extracting and processing those records. This program queries and reports any user-specified records and their attributes, writing each record type to a new csv file. 

# Dependencies
This program was created using C#, the QuickBooks Desktop SDK (QBSDK) from Intuit Developer, the CsvHelper package from Josh Close on GitHub, and the System.Web.Extensions library contained in .NET Framework. 

According to the QBSDK documentation, the following QuickBooks editions are supported: "just about every member of the QuickBooks Financial Software products family, for the U.S. (2002*—2018), Canada (2003*—2018), and the UK (2003*—2018). The SDK works with all U.S. editions of QuickBooks 2006 and later except for QuickBooks for the Mac."

QuickBooks must be running with the desired company file open in order for this application to function.

# To Run
In the QuickBooks-Proof-Of-Concept/bin/Release folder, edit the config.json file to reflect the desired query elements and their respective attributes. Each must exactly match the QuickBooks object name. Save this file. Open QuickBooks to the desired company file. Run the Quicken-Proof-Of-Concept.exe file, also contained within the Release folder. One csv file will be generated for each query, saved to the same location as the program, and the names of these files will be printed to the terminal by the program.

# Links
QBSDK: https://developer.intuit.com/docs/01_quickbooks_desktop/1_get_started/00_get_started <br>
CsvHelper: https://github.com/joshclose/csvhelper
