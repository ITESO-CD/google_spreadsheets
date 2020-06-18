
# -- --------------------------------------------------------------------------------------------------- -- #
# -- project: Google Spreadsheets conexion codes                                                         -- #
# -- script: main.py - A functional example of conectivity with google spreadsheets                      -- #
# -- author: FranciscoME                                                                                 -- #
# -- license: GPL-3.0                                                                                    -- #
# -- repository: https://github.com/ITESO-CD/google_spreadsheets                                         -- #
# -- --------------------------------------------------------------------------------------------------- -- #

import numpy as np
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials


# -- ----------------------------------------------------------- FUNCTION: Read/Write Google SpreadSheet -- #
# -- --------------------------------------------------------------------------------------------------- -- #

def f_google_ss(p_credentials, p_file, p_spreadsheet, p_option, p_position, p_data=None):
    """
    Function that receives a credential file (p_crendetials), and connects with google spreadsheets
    (through the Google API, and reads a specific file (p_file) and a specific spreadsheet (p_spreadsheet)

    Parameters
    ----------
    p_credentials : str : the file name and extension of the file with credentials
    p_file : str : the name of the google spreadsheet file to be read
    p_spreadsheet : str : the name of the sheet in the google spreadsheet file to be read
    p_option : str : 'read' or 'write' option
    p_data : pd.DataFrame : with the data to upload
    p_position : str : cell location to start writing the data

    Returns
    -------
    r_f_google_ss : str : dictionary with output elements

    Debugging
    ---------
    p_credentials = 'credentials/mt4-spreadsheet.json'
    p_file = 'Optimizacion_MetaQuotes'
    p_spreadsheet = 'Resultados_bt'
    p_option = 'write'
    p_data = df_f_data
    p_position = 'D4'
    """

    # list with scope urls
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # specify file with credentials (downloaded from google)
    creds = ServiceAccountCredentials.from_json_keyfile_name(p_credentials, scope)
    # initialize cliente with credentials
    client = gspread.authorize(creds)

    # -- if p_option == 'read'

    if p_option == 'read':
        # read the desired file and a specific spreadsheet
        backtest_file = client.open(p_file).worksheet(p_spreadsheet)
        # extract data of interest
        backtest_data = backtest_file.get_all_records()
        # results for the output
        r_f_google_ss = {'data': backtest_data}

        return r_f_google_ss

    # -- if p_option == 'write'

    else:
        # read the desired file and a specific spreadsheet
        backtest_file = client.open(p_file).worksheet(p_spreadsheet)
        # write entire dataframe
        r_f_google_ss = backtest_file.update(p_position, [p_data.columns.values.tolist()] +
                                             p_data.values.tolist())

        return r_f_google_ss
