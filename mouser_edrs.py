#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
This script uses the Firefox geckodriver to automatically fill in a University of Camrbidge Dept of Engineering (EDRS)
requisition with items from a Mouser Electronics Shopping cart exported as Excel file.

Download the geckodriver from https://github.com/mozilla/geckodriver/releases and save it in same directory as this script

This script can be easily extended to other companies or to use Chrome instead of Firefox.


Created on Thu May 23 08:22:26 2024

@author: Tobias E. Naegele, ten26
"""

import sys
from tkinter import filedialog, messagebox
import tkinter as tk
import pandas as pd
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium import webdriver

GECODRIVER_PATH = r'./geckodriver'

# associate column headers in mouser excel file with categories in EDRS system, don't need to edit this
MOUSER_MATCH_FIELDS = {
    'Mouser No': 'partno',
    'Order Qty.': 'qty',
    'Description ': 'description',
    'Price (GBP)': 'unitprice'}

def get_field_numbers():
    """
    identify the id of each row in edrs, sometimes edrs counts up, sometimes it gives them numbers starting with "n"
    this catches all cases

    Returns
    -------
    field_numbers : list
        List of all row ids
    """
    field_numbers = []
    ret = driver.find_elements(By.CSS_SELECTOR, "input[id*='description-']")
    for i in ret:
        a, num = i.get_attribute(By.NAME).split('-')
        field_numbers.append(num)
    return field_numbers


def parse_mouser_quote(path):
    """
    Parses a mouser quotation excel file exported from the shopping cart

    Parameters
    ----------
    path : str
        Path to the mouser quoate excel file.

    Returns
    -------
    df : pandas dataframe
        Pandas dataframe containing the shopping cart items.

    """
    df = pd.read_excel(path, skiprows=8, header=0)
    df = df.rename(columns={'Unnamed: 0': 'item_no'})
    df = df[df['Mouser No'].notna()]
    df = df.astype({'Order Qty.': int})
    df['Price (GBP)'] = df['Price (GBP)'].replace({'\\£': '', ',': ''}, regex=True).astype(
        float)  # see https://pbpython.com/currency-cleanup.html
    return df


def extend_edrs_lines(length):
    """
    Extends length of the edrs requisition until length is reached

    Parameters
    ----------
    length : int
        Desired requisition length

    Returns
    -------
    None.

    """
    field_numbers = get_field_numbers()
    while len(field_numbers) < length:
        lineadd = driver.find_element(By.NAME, "addmorelines")
        lineadd.click()
        time.sleep(0.1)
        field_numbers = get_field_numbers()
    return 0


def set_category(category):
    """
    Set category of all EDRS requisition lines to 'category'

    Parameters
    ----------
    category : str
        Category to be set on all lines. Example: 'LZ

    Returns
    -------
    None.

    """
    field_numbers = get_field_numbers()
    for i in field_numbers:
        ret = driver.find_element(
            By.CSS_SELECTOR,
            f"input[id='categorycode-{i}']")
        ret.send_keys(category)
    return 0


def fill_fields(data, match_fields):
    """
    Fills fields of EDRS requisition with items from data. Use match_fields to match data column headers with EDRS categories

    Parameters
    ----------
    data : pandas.DataFrame
        Pandas dataframe containing items to be added to EDRS order.
    match_fields : dict
        Dictionary which matches column names of data with EDRS categories. EDRS categories are ['qty','partno','description','unitprice','categorycode'].
        Example: match_fields = {'Mouser No':'partno','Order Qty.':'qty','Description ':'description','Price (GBP)':'unitprice'}

    Returns
    -------
    None.

    """
    field_numbers = get_field_numbers()
    for idx, i in enumerate(field_numbers):
        for df_column, edrs_element in match_fields.items():
            ret = driver.find_element(
                By.CSS_SELECTOR, f"input[id='{edrs_element}-{i}']")
            ret.send_keys(str(df[df_column][idx]))
    return 0


def select_file():
    """
    Opens dialog to ask for file

    Returns
    -------
    file_path : str

    """
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    root.destroy()
    return file_path


def ask_to_continue():
    """
    Opens dialog to ask to continue

    Returns
    -------
    file_path : str

    """
    root = tk.Tk()
    root.withdraw()
    ret = messagebox.askokcancel(
        "User action required",
        "Log into EDRS, create a new requisiotion and navigate to Items page. Press ok when ready")
    root.update()
    root.destroy()
    if not ret:
        print('Aborted.')
        sys.exit()
    return 0

###################


mouser_quoate_path = select_file()

# set up selenium driver
service = Service(executable_path=GECODRIVER_PATH)
options = webdriver.FirefoxOptions()
driver = webdriver.Firefox(service=service, options=options)

# navigate to EDRS webpage
driver.get('https://edrs.eng.cam.ac.uk/req/')

ask_to_continue()
# input('Log into EDRS, create a new requisiotion and navigate to Items page. Press return when ready')

# parse quote and extend requisition length
df = parse_mouser_quote(mouser_quoate_path)
extend_edrs_lines(len(df))

# fill fields
fill_fields(df, MOUSER_MATCH_FIELDS)

# set category to LZ
set_category('LZ')

print('Done.')
