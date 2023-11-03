import time
import math
import numpy as np
import pandas as pd
import win32com.client as win32
import matplotlib.pyplot as plt

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)


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


def plot_matplotlib(sheet_source, categories, range_yaxis, range_excel, titel):
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
    categories = range_to_array(sheet_source, categories).flatten()
    y_axis = range_to_array(sheet_source, range_yaxis).flatten()
    excel_data = range_to_array(sheet_source, range_excel)
    fig, ax = plt.subplots()

    for i in range(excel_data.shape[1]):
        color = plt.get_cmap("summer", len(categories))(i)
        ax.plot(y_axis, excel_data[:, i], label=f'{categories[i]}', color=color)
    ax.set_title(titel, weight='bold')
    ax.legend(loc='upper left')
    plt.show()


def plot_tabelle(sheet_source, range_excel, sheet_destination, position, chart_type, titel):
    """
    Generates an excel chart from a range of data in a fixed position.

    Parameters
    ----------
    sheet_source: string
    Represents the excel sheet where the data is located.

    range_excel: string
    Represents the range of data in an excel format (example: "A2:E32").

    sheet_destination: string
    Represents the excel sheet where the chart is generated.

    position: string
    Represents the location of the generated chart.

    chart_type: int
    Specifies the chart type by following the XlChartType-Enumeration (Including index and headers)

    Returns
    ----------
    No return specified

    Examples
    ----------
    plot_tabelle(sheet_source="Test Sheet", range_excel="A2:E32", sheet_destination="Test Sheet", position="Q25",
                 chart_type=72)

    """
    excel = win32.GetActiveObject("Excel.Application")
    wb = excel.Workbooks("Mainova_VersionAOF.xlsm")
    sheet_source = wb.Sheets(sheet_source)
    sheet_destination = wb.Sheets(sheet_destination)
    chart_location = sheet_destination.Range(position)
    chart = sheet_destination.Shapes.AddChart2(8, chart_type, chart_location.Left, chart_location.Top, 595, 270)
    # AddChart2(Diagrammformat, Typ des Diagramms, Position des linken Rands des Diagramms, Position des oberen Rands
    # des Diagramms, Breite, Höhe)
    chart.Chart.SetSourceData(sheet_source.Range(range_excel), 2)
    chart.Chart.ChartTitle.Text = titel
    chart.Chart.ChartTitle.Font.Bold = True
    chart.Chart.HasLegend = True


def range_to_array(excel_sheet, excel_value):
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
    wb = excel.Workbooks("Mainova_VersionAOF.xlsm")
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


def jahresverbrauch_tabelle(destatis_datei, dashboard_datei):
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

    jahresverbrauch = dashboard_datei.iloc[3, 2]

    if math.isnan(jahresverbrauch):
        segment = dashboard_datei.iloc[5, 2]
        if segment == "Haushalte":
            erdgas_verbrauch = range_to_array("destatis_Daten_Kopie", "B7:D7")
            strom_verbrauch = range_to_array("destatis_Daten_Kopie", "L9:N9")

        else:
            erdgas_verbrauch = range_to_array("destatis_Daten_Kopie", "G10:I10")
            strom_verbrauch = range_to_array("destatis_Daten_Kopie", "Q11:S11")

    else:
        erdgas_verbrauch = np.array(erdgas_verbrauch.loc[
                                        (erdgas_verbrauch.index.get_level_values("min_value") <= jahresverbrauch) & (
                                                erdgas_verbrauch.index.get_level_values(
                                                    "max_value") >= jahresverbrauch)])
        strom_verbrauch = np.array(strom_verbrauch.loc[
                                       (strom_verbrauch.index.get_level_values("min_value") <= jahresverbrauch) & (
                                               strom_verbrauch.index.get_level_values("max_value") >= jahresverbrauch)])
    return erdgas_verbrauch, strom_verbrauch


def jahresverbrauch_berechnung(erdgas_verbrauch, strom_verbrauch):
    excel = win32.GetActiveObject("Excel.Application")
    wb = excel.Workbooks("Mainova_VersionAOF.xlsm")
    preisannahmen = wb.Sheets("Preisannahmen_Kopie")
    preisannahmen.Cells(4, 15).Value = erdgas_verbrauch[0, 1] * 100
    preisannahmen.Cells(4, 16).Value = (erdgas_verbrauch[0, 2] * 100) - (preisannahmen.Cells(5, 21).Value +
                                                                         preisannahmen.Cells(6, 21).Value)
    preisannahmen.Cells(4, 17).Value = strom_verbrauch[0, 1] * 100
    preisannahmen.Cells(4, 18).Value = strom_verbrauch[0, 2] * 100 - preisannahmen.Cells(4, 21).Value

    plot_tabelle(sheet_source="Preisannahmen_Kopie", range_excel="N2:R54", sheet_destination="Dashboard",
                 position="G9", chart_type=72, titel="Jährliche Kostenannahmen")
    plot_matplotlib(sheet_source="Preisannahmen_Kopie", categories="O2:R2", range_yaxis="A4:A54", range_excel="O4:R54",
                    titel="Jährliche Kostenannahmen")
    return


def kessel(klasse, parameter_datei, leistungs_datei):
    print("Kessel")
    if klasse == 1:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Kessel Gas":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 10), (position_tech[1] + 3):]

    elif klasse == 2:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Kessel Gas":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 10), (position_tech[1] + 3):]

    else:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Kessel Pellet":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 10), (position_tech[1] + 3):]
    wgk_berechnung = pd.concat([pd.DataFrame([leistungs_datei], columns=tech_datei.columns), tech_datei], ignore_index=True)
    print("TECHNOLOGIE DATEI", wgk_berechnung)
    # EXAMPLE wgk gesamt = (df.iloc[1] / df.iloc[0]) * df.iloc[2] + df.iloc[3]
    wgk_berechnung = (0.08 / (1500 * tech_datei.iloc[0])) + tech_datei.iloc[9]
    return tech_datei


def wp(klasse, parameter_datei):
    print("Wärmepumpe")
    if klasse == 1:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Wärmepumpe Luft-Wasser <75 kWh/m²a":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 11), (position_tech[1] + 3):]
    elif klasse == 2:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Wärmepumpe Luft-Wasser 75 - 130 kWh/m²a":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 11), (position_tech[1] + 3):]
    elif klasse == 3:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Wärmepumpe Luft-Wasser >130 kWh/m²a":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 11), (position_tech[1] + 3):]
    elif klasse == 4:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Wärmepumpe Sole-Wasser <75 kWh/m²a":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 13), (position_tech[1] + 3):]

    elif klasse == 5:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Wärmepumpe Sole-Wasser <75 kWh/m²a":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 13), (position_tech[1] + 3):]
    elif klasse == 6:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Wärmepumpe Sole-Wasser 75 - 130 kWh/m²a":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 13), (position_tech[1] + 3):]

    elif klasse == 7:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Wärmepumpe Sole-Wasser 75 - 130 kWh/m²a":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 13), (position_tech[1] + 3):]

    elif klasse == 8:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Wärmepumpe Sole-Wasser >130 kWh/m²a":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 13), (position_tech[1] + 3):]

    else:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Wärmepumpe Sole-Wasser >130 kWh/m²a":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 13), (position_tech[1] + 3):]
    return tech_datei


def wp_gk(klasse, parameter_datei):
    print("Wärmepumpe + GasKessel")
    for row_uno in range(parameter_datei.shape[0]):
        for col_uno in range(parameter_datei.shape[1]):
            if parameter_datei.iloc[row_uno, col_uno] == "Hybrid Luft-Wasser Wärmepumpe + Kessel Gas (Gaskessel, " \
                                                         "Wärmepumpe)":
                position_tech = (row_uno, col_uno)

    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 25), (position_tech[1] + 3):]

    if klasse == 1:
        klasse = 1
    else:
        klasse = 0

    wgk_tech = klasse * 1
    print("Position Tech", position_tech)
    print(tech_datei)

    return wgk_tech


def bhkw_gk(klasse, parameter_datei):
    print("Blockheizkraftwerke")
    if klasse == 1:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid BHKW + Kessel Gas 25% Eigenstromquote (" \
                                                             "Gaskessel, BHKW)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 29), (position_tech[1] + 3):]
    elif klasse == 2:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid BHKW + Kessel Gas 25% Eigenstromquote (" \
                                                             "Gaskessel, BHKW)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 29), (position_tech[1] + 3):]
    elif klasse == 3:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid BHKW + Kessel Gas 50% Eigenstromquote (" \
                                                             "Gaskessel, BHKW)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 29), (position_tech[1] + 3):]
    else:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid BHKW + Kessel Gas 50% Eigenstromquote (" \
                                                             "Gaskessel, BHKW)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 29), (position_tech[1] + 3):]

    wgk_tech = klasse * 1
    print("Position Tech", position_tech)
    print(tech_datei)

    return wgk_tech


def st_gk(klasse, parameter_datei):
    print("Solarthermie + Gas Kessel")
    if klasse == 1:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid Solarthermie + Kessel Gas 20 % ST (Gaskessel, " \
                                                             "Solarthermie)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 21), (position_tech[1] + 3):]
    elif klasse == 2:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid Solarthermie + Kessel Gas 20 % ST (Gaskessel, " \
                                                             "Solarthermie)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 21), (position_tech[1] + 3):]
    elif klasse == 3:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid Solarthermie + Kessel Gas 35 % ST (Gaskessel, " \
                                                             "Solarthermie)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 21), (position_tech[1] + 3):]
    else:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid Solarthermie + Kessel Gas 35 % ST (Gaskessel, " \
                                                             "Solarthermie)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 21), (position_tech[1] + 3):]

    wgk_tech = klasse * 1
    print("Position Tech", position_tech)
    print(tech_datei)
    return wgk_tech


def st_lwwp(klasse, parameter_datei):
    print("Solarthermie + Luft-Wasser Wärmepumpe")
    if klasse == 1:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid Solarthermie + Luft-Wasser Wärmepumpe 20 % ST (Wärmepumpe, Solarthermie)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 22), (position_tech[1] + 3):]
    else:
        for row_uno in range(parameter_datei.shape[0]):
            for col_uno in range(parameter_datei.shape[1]):
                if parameter_datei.iloc[row_uno, col_uno] == "Hybrid Solarthermie + Luft-Wasser Wärmepumpe 35 % ST (Wärmepumpe, Solarthermie)":
                    position_tech = (row_uno, col_uno)
                    tech_datei = parameter_datei.iloc[position_tech[0]:(position_tech[0] + 22), (position_tech[1] + 3):]

    wgk_tech = klasse * 1
    print("Position Tech", position_tech)
    print(tech_datei)
    return wgk_tech


def warmegestehungskosten_berechnung(parameter_datei):
    excel = win32.GetActiveObject("Excel.Application")
    wb = excel.Workbooks("Mainova_VersionAOF.xlsm")
    dashboard_datei = wb.Sheets("Dashboard")
    technologie_value = dashboard_datei.Cells(5, 3).Value
    print("READ TECHNOLOGIE INPUT", technologie_value)
    leistungs_klasse = dashboard_datei.Cells(7, 3).Value
    print("Technologie DATEI: ", technologie_value)
    print("Leistungsklasse: ", leistungs_klasse)
    investitionsjahr = dashboard_datei.Cells(7, 3).Value

    leistungs_datei = parameter_datei.iloc[2, 3:]
    technologie_dict = {"Kessel Gas": (kessel, 1, leistungs_datei),
                        "Kessel Gas (Anschluss vorhanden)": (kessel, 2),
                        "Kessel Pellet": (kessel, 3),
                        "Wärmepumpe Luft-Wasser <75 kWh/m²a": (wp, 1),
                        "Wärmepumpe Luft-Wasser 75 - 130 kWh/m²a": (wp, 2),
                        "Wärmepumpe Luft-Wasser >130 kWh/m²a": (wp, 3),
                        "Wärmepumpe Sole-Wasser <75 kWh/m²a": (wp, 4),
                        "Wärmepumpe Sole-Wasser <75 kWh/m²a (Wärmequelle vorhanden)": (wp, 5),
                        "Wärmepumpe Sole-Wasser 75 - 130 kWh/m²a": (wp, 6),
                        "Wärmepumpe Sole-Wasser 75 - 130 kWh/m²a (Wärmequelle vorhanden)": (wp, 7),
                        "Wärmepumpe Sole-Wasser >130 kWh/m²a": (wp, 8),
                        "Wärmepumpe Sole-Wasser >130 kWh/m²a (Wärmequelle vorhanden)": (wp, 9),
                        "Hybrid Luft-Wasser Wärmepumpe + Kessel Gas": (wp_gk, 1),
                        "Hybrid Luft-Wasser Wärmepumpe + Kessel Gas (Anschluss vorhanden)": (wp_gk, 2),
                        "Hybrid BHKW + Kessel Gas 25% Eigenstromquote": (bhkw_gk, 1),
                        "Hybrid BHKW + Kessel Gas 25 % Eigenstromquote(Anschluss vorhanden)": (bhkw_gk, 2),
                        "Hybrid BHKW + Kessel Gas 50% Eigenstromquote": (bhkw_gk, 3),
                        "Hybrid BHKW + Kessel Gas 50% Eigenstromquote (Anschluss vorhanden)": (bhkw_gk, 4),
                        "Hybrid Solarthermie + Kessel Gas 20 % ST": (st_gk, 1),
                        "Hybrid Solarthermie + Kessel Gas 20 % ST (Anschluss vorhanden)": (st_gk, 2),
                        "Hybrid Solarthermie + Kessel Gas 35 % ST": (st_gk, 3),
                        "Hybrid Solarthermie + Kessel Gas 35 % ST (Anschluss vorhanden)": (st_gk, 4),
                        "Hybrid Solarthermie + Luft-Wasser Wärmepumpe 20 % ST": (st_lwwp, 1),
                        "Hybrid Solarthermie + Luft-Wasser Wärmepumpe 35 % ST": (st_lwwp, 2)
                        }

    if technologie_value in technologie_dict:
        wgk, klasse, leistungs_klasse = technologie_dict[technologie_value]
        print("Parameter Klasse: ", klasse)
        print("WGK results", wgk(klasse, parameter_datei, leistungs_klasse))

    else:
        print("Invalid Technologie input")

    if leistungs_klasse in leistungs_dict:
        leistungs_value = leistungs_dict[leistungs_klasse]
        print("Ergebnis Position", leistungs_value)
    else:
        print("Invalid Leistungsklasse input")
    return


if __name__ == "__main__":

    destatis_datei = pd.read_excel(io="/Users/aochoa/PycharmProjects/Fraunhofer/Mainova_VersionAOF.xlsm",
                                   sheet_name="destatis_Daten_Kopie", header=None)
    dashboard_datei = pd.read_excel(io="/Users/aochoa/PycharmProjects/Fraunhofer/Mainova_VersionAOF.xlsm",
                                    sheet_name="Dashboard", header=None)
    parameter_datei = pd.read_excel(io="/Users/aochoa/PycharmProjects/Fraunhofer/Mainova_VersionAOF.xlsm",
                                    sheet_name="Inputparameter_Kopie", header=None)
    """
    plot_matplotlib(sheet_source="Test Sheet", range_xaxis="B2:E2", range_yaxis="A4:A32", range_excel="B4:E32",
                    sheet_destination="Test Sheet", position="Q1")

    plot_tabelle(sheet_source="Test Sheet", range_excel="A2:E32", sheet_destination="Test Sheet", position="Q25",
                 chart_type=72)
    
    erdgas_verbrauch, strom_verbrauch = jahresverbrauch_tabelle(destatis_datei, dashboard_datei)
    jahresverbrauch_berechnung(erdgas_verbrauch, strom_verbrauch)
    """
    warmegestehungskosten_berechnung(parameter_datei)