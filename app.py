# -*- coding: utf-8 -*-

"""
Calculate heat processing cost for different scenarios.

SPDX-FileCopyrightText: Uwe Krien <uwe.krien@ifam.fraunhofer.de>
SPDX-FileCopyrightText: Esmail Ansari <esmail.ansari@ifam.fraunhofer.de

SPDX-License-Identifier: MIT
"""
import math
import numpy as np
import win32com.client as win32
import matplotlib.pyplot as plt
import pandas as pd
import subprocess
import datetime
import logging
import logging.config
import os
import sys
import warnings
from tkinter import filedialog as fd
from scenario import Scenario

subprocess.call([sys.executable, '-m', 'pip', 'install', "numpy"])
subprocess.call([sys.executable, '-m', 'pip', 'install', "pandas"])
subprocess.call([sys.executable, '-m', 'pip', 'install', "matplotlib"])
subprocess.call([sys.executable, '-m', 'pip', 'install', "pywin32"])

warnings.simplefilter(action="ignore", category=pd.errors.PerformanceWarning)
logging.basicConfig(
    stream=sys.stdout,
    level=logging.INFO,
    format="%(asctime)s, %(levelname)s %(name)s %(message)s",
)


def my_app(file=None):
    """
    Main application file.
    """
    logging.info("************ Running WGKPy *****************")
    n = 0
    logging.debug(
        "Working directory: {0}".format(os.path.dirname(sys.argv[0]))
    )
    for a in sys.argv:
        n += 1
        logging.debug("Argument {0}: {1}".format(n, a))
    logging.debug("{0}".format(sys.argv))
    if len(sys.argv) > 1:
        file = sys.argv[1]
    logging.debug("Given filename: {0}".format(file))

    if file is None:
        file = fd.askopenfilename()
        logging.debug("Chosen filename: {0}".format(file))

    start = datetime.datetime.now()
    logging.info("Start: {0}".format(os.path.basename(file)))
    logging.info("Path: %s" % file)

    # sc = Scenario(file, flh=[1500, 2000], years=[2028, 2033])
    sc = Scenario(file, sheet_keyword="Inputparameter")
    sc.create_result_table()
    logging.info("Results - %s" % str(datetime.datetime.now() - start))
    sc.results_to_xlsx()
    logging.info("Done! - %s" % str(datetime.datetime.now() - start))
    return file


def jahresverbrauch_tabelle(destatis_datei, dashboard_datei, file):
    """
    Uses the information displayed within the dashboard sheet in order to return the
    Parameters
    ----------
    :param destatis_datei: 'pandas.core.frame.DataFrame'
    Represents the data located in the destatis sheet.

    :param dasboard_datei: 'pandas.core.frame.DataFrame'
    Represents the data located in the dashboard sheet.
    ----------

    :return:
    """
    for row_uno in range(destatis_datei.shape[0]):
        for col_uno in range(destatis_datei.shape[1]):
            if destatis_datei.iloc[row_uno, col_uno] == "Jahresverbrauchsklassen (Erdgaspreise)":
                position_uno = (row_uno, col_uno)
    for row_dos in range(destatis_datei.shape[0]):
        for col_dos in range(destatis_datei.shape[1]):
            if destatis_datei.iloc[row_dos, col_dos] == "Jahresverbrauchsklassen (Strompreise)":
                position_dos = (row_dos, col_dos)

    erdgas_verbrauch = destatis_datei.iloc[position_uno[0]:(position_uno[0] + 10),
                       position_uno[1]:(position_uno[1] + 5)]
    erdgas_verbrauch.set_index([erdgas_verbrauch.columns[0], erdgas_verbrauch.columns[1]], inplace=True)
    erdgas_verbrauch.index.names = ['min_value', 'max_value']
    erdgas_verbrauch = erdgas_verbrauch.drop(index=erdgas_verbrauch.index[:2])
    erdgas_verbrauch = erdgas_verbrauch.rename(index={'mehr': 999999999})

    strom_verbrauch = destatis_datei.iloc[position_dos[0]:(position_dos[0] + 13), position_dos[1]:(position_dos[1] + 5)]
    strom_verbrauch.set_index([strom_verbrauch.columns[0], strom_verbrauch.columns[1]], inplace=True)
    strom_verbrauch.index.names = ['min_value', 'max_value']
    strom_verbrauch = strom_verbrauch.drop(index=strom_verbrauch.index[:2])
    strom_verbrauch = strom_verbrauch.rename(index={'mehr': 999999999})

    jahresverbrauch = dashboard_datei.iloc[3, 10]
    if math.isnan(jahresverbrauch):
        segment = dashboard_datei.iloc[7, 10]
        if segment == "Haushalte":
            erdgas_verbrauch = range_to_array("destatis_Daten_Kopie", "B7:D7", file)
            strom_verbrauch = range_to_array("destatis_Daten_Kopie", "L9:N9", file)

        else:
            erdgas_verbrauch = range_to_array("destatis_Daten_Kopie", "G10:I10", file)
            strom_verbrauch = range_to_array("destatis_Daten_Kopie", "Q11:S11", file)

    else:
        erdgas_verbrauch = np.array(erdgas_verbrauch.loc[
                                        (erdgas_verbrauch.index.get_level_values("min_value") <= jahresverbrauch) & (
                                            erdgas_verbrauch.index.get_level_values(
                                                "max_value") >= jahresverbrauch)])
        strom_verbrauch = np.array(strom_verbrauch.loc[
                                       (strom_verbrauch.index.get_level_values("min_value") <= jahresverbrauch) & (
                                           strom_verbrauch.index.get_level_values("max_value") >= jahresverbrauch)])
    return erdgas_verbrauch, strom_verbrauch


def range_to_array(excel_sheet, excel_value, file):
    """
    Converts an Excel range into a numpy array. It can also return the coordinates of a single cell value if the input
    is not an Excel range.

    Parameters
    ----------
    excel_sheet: string
    Represents the Excel sheet where the range is located.

    excel_value: string
    represents the range of the data.

    ----------
    return: numpy array
    Returns the values of the cells in the format of a numpy array

    Examples
    ----------
    range_to_array("Sheet 1", "AA4:AC32")
    """
    excel = win32.GetActiveObject("Excel.Application")
    wb = excel.Workbooks(file)
    excel_sheet = wb.Sheets(excel_sheet)
    split_strings = excel_value.split(":")
    if len(split_strings) != 2:
        # Gives coordinate of excel_value if there is no range
        empty_array = (excel_konverter(''.join(filter(str.isalpha, split_strings[0]))),
                       int(''.join(filter(str.isdigit, split_strings[0]))))
    else:
        # Saves the data found within the Excel range into a numpy array
        l1 = excel_konverter(''.join(filter(str.isalpha, split_strings[0])))
        l2 = excel_konverter(''.join(filter(str.isalpha, split_strings[1])))
        n1 = int(''.join(filter(str.isdigit, split_strings[0])))
        n2 = int(''.join(filter(str.isdigit, split_strings[1])))
        num_col = l2 - l1 + 1
        num_row = n2 - n1 + 1
        empty_array = np.empty((num_row, num_col), dtype=object)
        for row in range(num_row):
            for col in range(num_col):
                try:
                    empty_array[row, col] = float(excel_sheet.Cells(n1 + row, l1 + col).Value)
                except ValueError:
                    empty_array[row, col] = str(excel_sheet.Cells(n1 + row, l1 + col).Value)

    return empty_array


def excel_konverter(value):
    """
    Converts an Excel format base26 column string to a number and viceversa depending on the parameter type given.

    Parameters
    ----------
    column: string or int
    Represents the column name in the Excel format (string) or its mathematical location (int).

    ----------
    return: int or string
    Represents the calculated mathematical location (int) or the column name in the Excel format (string).

    Examples
    ----------

    >>> excel_konverter("AA")
    27

    >>> excel_konverter(122)
    DR
    """
    if isinstance(value, str):
        """ Convert base26 column string to number. """
        expn = 0
        result = 0
        for char in reversed(value):
            result += (ord(char) - ord('A') + 1) * (26 ** expn)
            expn += 1

        return result

    elif isinstance(value, int):
        result = ""
        """ Convert number to base26 column string. """
        while value > 0:
            residuo = (value - 1) % 26
            result = chr(ord('A') + residuo) + result
            value = (value - 1) // 26

        return result

    else:
        raise ValueError("Input must be either a string or an integer")


def jahresverbrauch_berechnung(erdgas_verbrauch, strom_verbrauch, file):
    excel = win32.GetActiveObject("Excel.Application")
    wb = excel.Workbooks(file)
    preisannahmen = wb.Sheets("Preisannahmen")
    preisannahmen.Cells(4, 15).Value = erdgas_verbrauch[0, 1] * 100
    preisannahmen.Cells(4, 16).Value = (erdgas_verbrauch[0, 2] * 100) - (preisannahmen.Cells(5, 21).Value + 0.4575)
    preisannahmen.Cells(4, 17).Value = strom_verbrauch[0, 1] * 100
    preisannahmen.Cells(4, 18).Value = strom_verbrauch[0, 2] * 100 - preisannahmen.Cells(4, 21).Value

    plot_matplotlib(sheet_source="Preisannahmen", categories="O2:R2", range_yaxis="A4:A50", range_excel="O4:R50",
                    titel="Jährliche Kostenannahmen", file=file)
    return


def plot_matplotlib(sheet_source, categories, range_yaxis, range_excel, titel, file):
    """
    Generates a matplotlib line chart from a range of data in a fixed position.

    Parameters
    ----------
    sheet_source: string
    Represents the Excel sheet where the data is located.

    range_xaxis: string
    represents the range of headers of the data.

    range_yaxis: string
    represents the range of index of the data.

    range_excel: string
    Represents only the range of data to be plotted

    sheet_destination: string
    Represents the excel sheet where the chart is generated.

    position: string
    Represents the location of the generated chart.

    Returns
    ----------
    No return specified

    Examples
    ----------
    plot_matplotlib(sheet_source="Test Sheet", range_xaxis="B2:E2", range_yaxis="A4:A32",range_excel="B4:E32",
                    sheet_destination="Test Sheet", position="Q1")
    """
    categories = range_to_array(sheet_source, categories, file).flatten()
    y_axis = range_to_array(sheet_source, range_yaxis, file).flatten()
    excel_data = range_to_array(sheet_source, range_excel, file)
    fig, ax = plt.subplots()

    for i in range(excel_data.shape[1]):
        color = plt.get_cmap("summer", len(categories))(i)
        ax.plot(y_axis, excel_data[:, i], label=f'{categories[i]}', color=color)
    ax.set_title(titel, weight='bold')
    ax.legend(loc='upper left')
    plt.show()


if __name__ == "__main__":
    file = my_app()
    destatis_datei = pd.read_excel(io=file, sheet_name="destatis_Daten_Kopie", header=None)
    dashboard_datei = pd.read_excel(io=file, sheet_name="Dashboard", header=None)
    path, file = os.path.split(file)
    erdgas_verbrauch, strom_verbrauch = jahresverbrauch_tabelle(destatis_datei, dashboard_datei, file)
    jahresverbrauch_berechnung(erdgas_verbrauch, strom_verbrauch, file)

