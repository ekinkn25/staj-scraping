### Staj Scrape
This is a python based web scraping project that extracts the website URLs of individual companies from various technology parks directory pages.

Please note: Since every website has a different design, this script might not work for other domains. It is currently optimized only for Bilkent CyberPark and ODTÜ Teknokent pages.

Prerequisites:
* Python (tested with Python 3.12)
* beautifulsoup4
* pandas
* openpyxl


You can install all dependencies via your terminal or command prompt using pip:

`pip install requests beautifulsoup4 pandas openpyxl`

#### Usage : 
* Input: Create or edit linkler.txt in the root directory of the project. Paste the target technology park directory URLs into this file, ensuring each URL is on a new line.

* Execution: Run the script through your terminal `python code.py`

* Processing: Script reads the URLs from the text file one by one. It automatically filters out empty entries and normalizes missing links.

* Dynamic Export: Gathered data is exported to an Excel file (.xlsx). If company has no corrresponding URL on web site it still would be shown at the Excel file but would written "Link bulunamadı" in URL coulumn.
