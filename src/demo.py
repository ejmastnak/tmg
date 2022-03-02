import pandas as pd
import numpy as np
import os

import tmg, constants

"""
A few convenience functions for calling the
TMG-parameter-calculation functions in `tmg_params.py`.
"""

def xlsx_to_pandas_df(xlsx_file, max_signal_rows=constants.TMG_MAX_ROWS):
    """
    Utility function for reading the measurements in a TMG-formatted Excel file
    into a Pandas dataframe.
    Drops the following information from the Excel file:
    - The first column  (which does not contain TMG signal data)
    - The first constants.DATA_START_ROW rows (contain metadata but no signal)

    Parameters
    ----------
    xlsx_file : str
        Full path to a TMG-formatted Excel measurement file
    max_signal_rows : int
        Number of rows (i.e. data points, i.e. milliseconds assuming
        1 kHz sampling) of inputted TMG signal to analyze, since most
        relevant information occurs in the first few hundred milliseconds only.

    Returns
    -------
    df : DataFrame
        Pandas dataframe equivalent of Excel file's measurement
    
    """
    return pd.read_excel(xlsx_file, engine='openpyxl', header=None,
            skiprows=constants.TMG_DATA_START_ROW,
            nrows=max_signal_rows).drop(columns=[0])


def measurement_analysis_wrapper(measurement_num=0, xlsx_file = "../sample-data/EM.xlsx"):
    """
    Wrapper method for computing TMG parameters of a single TMG 
    measurements in a standard TMG-formatted Excel file.

    Parameters
    ----------
    measurement_num : int
        Zero-based measurement number for which to calculate parameters.
    
    """
    # Read measurements from Excel file into a Pandas DataFrame
    df = xlsx_to_pandas_df(xlsx_file)

    # Extract 1D TMG signal as Pandas Series
    tmg_signal = df.iloc[:, measurement_num]

    # Compute TMG parameters
    params = tmg.get_params_of_tmg_signal(tmg_signal.to_numpy())

    # Print params in human-readable format
    param_names = constants.TMG_PARAM_NAMES
    for i, param in enumerate(params):
        print("{} {:.2f}".format(param_names[i], param))


def file_analysis_wrapper(xlsx_input_file = "../sample-data/EM.xlsx",
        output_file = "../sample-output/EM-params.csv"):
    """
    Wrapper method for computing TMG parameters of all TMG 
    measurements in a standard TMG-formatted Excel file.

    """
    # Read measurements from Excel file into a Pandas DataFramee
    df = xlsx_to_pandas_df(xlsx_input_file)

    measurement_names = []
    param_names = constants.TMG_PARAM_NAMES

    # First add parameters to a list, then create a DataFrame from the list
    param_list = []

    # Loop through each measurement number and TMG signal in Excel file
    for (m, tmg_signal) in df.iteritems():
        params = tmg.get_params_of_tmg_signal(tmg_signal.to_numpy())
        param_list.append(params)
        measurement_names.append("Measurement {}".format(m))

    param_df = pd.DataFrame(param_list).transpose()
    param_df.columns=measurement_names
    param_df.index=param_names
    param_df.to_csv(output_file)


def directory_analysis_wrapper(input_dir="../sample-data/",
        output_dir="../sample-output/"):
    """
    Wrapper method for computing TMG parameters for all
    TMG-formatted Excel files in a directory.

    """
    for xlsx_filename in sorted(os.listdir(input_dir)):
        if ".xlsx" in xlsx_filename and "$" not in xlsx_filename:
            output_filename = xlsx_filename.replace(".xlsx", "-params.csv")
            file_analysis_wrapper(xlsx_input_file=input_dir + xlsx_filename,
                    output_file=output_dir + output_filename)

    
if __name__ == "__main__":
    measurement_analysis_wrapper()
    file_analysis_wrapper()
    directory_analysis_wrapper()
