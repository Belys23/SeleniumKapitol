HEIMDALL - Client Search Software
Overview
HEIMDALL is a C# console application developed by Martin Belka to automate client searches on the Kapitol portal (https://portal.kapitol.cz/). It uses Selenium WebDriver to scrape client data based on birth numbers or phone numbers provided in an input Excel file (Kontakty.xlsx). The program outputs results to an Excel file (pecovatel_vysledky.xlsx) on the user's desktop, indicating whether a client is associated with a specific Kapitol service (excluding GOLD and AUTO services).
Features

Input Flexibility: Reads birth numbers or phone numbers from an Excel file (Kontakty.xlsx).
Search Options: Allows users to choose between searching by birth number or phone number.
Dynamic Row Selection: Users can specify the starting row and number of rows to process, with validation to prevent exceeding the input file's row count.
Web Scraping: Automates login and client search on the Kapitol portal using Selenium WebDriver with Chrome.
Result Filtering: Identifies clients associated with Kapitol services (excluding GOLD and AUTO) and marks them as "pozitivní" in the output.
Excel Output: Saves results to pecovatel_vysledky.xlsx on the desktop, including the searched number and result status.
Error Handling: Handles invalid inputs, disabled search buttons, and general exceptions with informative console messages.

Prerequisites

Operating System: Windows (due to ChromeDriver and Excel file handling).
Dependencies:
.NET Framework or .NET Core (compatible with the project setup).
EPPlus library for Excel file manipulation (non-commercial license).
Selenium.WebDriver and Selenium.WebDriver.ChromeDriver for web scraping.


Input File: An Excel file named Kontakty.xlsx in the same directory as the executable, with birth numbers in column 3 and phone numbers in column 1.
Chrome Browser: Installed and compatible with the ChromeDriver version used.
Kapitol Portal Credentials: Valid username and password for https://portal.kapitol.cz/.

Installation

Clone or Download the Project:
Obtain the project files from the repository or source.


Install Dependencies:
Use NuGet Package Manager to install:
EPPlus (Install-Package EPPlus -Version 5.8.7 or compatible).
Selenium.WebDriver (Install-Package Selenium.WebDriver).
Selenium.WebDriver.ChromeDriver (Install-Package Selenium.WebDriver.ChromeDriver).




Prepare Input File:
Place Kontakty.xlsx in the same directory as the executable, ensuring it has birth numbers in column 3 and phone numbers in column 1.


Set Up Credentials:
Replace "XXX" in the code (username and password fields) with valid Kapitol portal credentials.


Build the Project:
Open the solution in Visual Studio.
Build the solution to generate the executable.


Ensure ChromeDriver:
The ChromeDriver executable should match your installed Chrome version and be accessible in the project directory or PATH.



Usage

Run the Program:
Execute the compiled executable from the command line or Visual Studio.


Select Search Type:
Choose 1 for birth numbers or 2 for phone numbers when prompted.


Specify Rows:
Enter the starting row (1 to max rows in Kontakty.xlsx).
Enter the number of rows to process (automatically adjusted if exceeding max rows).


Processing:
The program logs into the Kapitol portal, navigates to the consultant page, and searches for each number.
Results are displayed in the console, showing the number and associated client data (if any).


Output:
Results are saved to pecovatel_vysledky.xlsx on the desktop, with columns for the searched number and result ("pozitivní" for Kapitol clients, excluding GOLD and AUTO).


Completion:
The program closes the browser and saves the output file. Press Enter to exit.



Example Workflow

Ensure Kontakty.xlsx is in the executable's directory.
Run the program and select 1 for birth numbers.
Enter starting row 1 and number of rows 10.
The program processes 10 birth numbers, outputs results to the console, and saves them to pecovatel_vysledky.xlsx.

Notes

Credentials: Hardcoded credentials (XXX) must be replaced with valid ones to avoid login failures.
Excel File: Ensure Kontakty.xlsx is correctly formatted to avoid runtime errors.
ChromeDriver: Update ChromeDriver if Chrome updates to maintain compatibility.
Performance: A 150ms delay is included after each search to ensure page loading. Adjust if needed for stability.
Error Handling: The program handles common errors (e.g., invalid inputs, disabled buttons) but may require manual intervention for network issues or portal changes.

Development

Language: C#.
Libraries:
EPPlus for Excel file reading/writing.
Selenium.WebDriver for browser automation.


Structure:
Reads input from Kontakty.xlsx using EPPlus.
Uses Selenium to automate Chrome browser interactions.
Outputs results to a new Excel file using EPPlus.


Limitations:
Requires manual credential input in the code.
Assumes a specific portal structure; updates to the Kapitol portal may break the scraper.



Author

Publisher: Morava Hemp
Developed as a client search automation tool.

License
This project uses a non-commercial license for EPPlus. Ensure compliance with EPPlus licensing terms. The project itself is not distributed under a specific license; contact the author for usage permissions.
