# Unit Availability Tracker for BTO HDB Flats

Got tired of manually checking availability so I made this script to automate the process.

Script uses selenium to scrape the HDB website for unit availability and generates an excel file with the data.

HDB only allows access to the unit info if you have a queue number for the BTO, so you will need to have a valid queue number and be logged in to the HDB website to use this script.

## Requirements
- Python 3.12 or higher (Havent tested on lower versions)
- Chrome
- ChromeDriver (Should be installed automatically by the script)

## Usage
1. Clone the repository

2. Install the required packages
   ```bash
   pip install -r requirements.txt
   ```
3. Make the necessary changes in the `generate_script_v2.py` file:
    - Set the `BTO_URL` variable to the URL of your BTO project.
    - Set the `BTO_UNIT_PREFIX` variable to the prefix of your BTO unit. This can be found in the URL of the BTO project page.
    Example:
    ```python
    BTO_URL = "https://homes.hdb.gov.sg/home/bto/details/2024-06_BTO_JSHFsjhfsjFSJHFsk"
    BTO_UNIT_PREFIX = "2024-06_BTO_"
    ```

4. You can optionally include a screenshot of the BTO layout by placing an image file named `layout.png` in the
   same directory as the script. The script will automatically include this image at the side of every generated
   sheet in the excel file. If you choose not have a layout image, the script will still run fine without it.

5. Run the script
   ```bash
   python generate_script_v2.py
   ```

6. When the browser opens, click on the sign in button, scan the QR code using your singpass app to sign into the HDB website. After allowing the sign in on your singpass app, do not click anything or close the browser window. The script will automatically navigate to the bto project page and scrape the data. After which, the script will generate an excel file with the data. If you click anything in the browser window during this process, it may interfere with the script's ability to scrape the data.

## Disclaimer
This script is not affiliated with or endorsed by the Housing & Development Board (HDB) of Singapore. Use at your own risk.

The script makes use of your HDB account credentials to access the unit availability information. Make sure you understand what the script is doing before running it.

As the script involves web scraping, it may break if the HDB website changes its structure or if there are any changes to the login process.
