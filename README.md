Excel API Processor

A simple ASP.NET Core MVC application that allows users to:
Upload an Excel file containing part numbers.
Provide an API URL to fetch part details dynamically.
Populate the Excel file with API data in the same columns.
Download the processed Excel file with proper formatting.

Features

Supports .xlsx, .xlsm, .xltx, .xltm files.
Dynamic mapping of columns – only fills columns that exist in the uploaded Excel.
Highlights missing API data in light pink.
Headers are bold with background color, columns auto-sized, and borders applied.
Clean, responsive Upload and Download pages.
Download page shows a button to download processed file and a link back to Upload page.
