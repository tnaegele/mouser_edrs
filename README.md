# mouser_edrs
Python script to create University of Cambridge Department of Engineering EDRS requisition from Mouser electronics quote using Selenium and Firefox.
This can easily be extended to accommodate other suppliers or to use Chromium instead of Firefox.

## Usage
1. Export your Mouser shopping cart (https://www.mouser.co.uk/Cart/) as Excel file
2. Start the script and choose the shopping cart Excel file
3. The script will open the EDRS webpage. Login with your Uni of Cambridge credentials, create a new requisition, choose the supplier and naviagte to the 'Items' page
4. Press ok in the dialog to confirm that you have done the steps above
