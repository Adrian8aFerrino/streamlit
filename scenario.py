# -*- coding: utf-8 -*-

"""
File handling.

SPDX-FileCopyrightText: Uwe Krien <uwe.krien@ifam.fraunhofer.de>

SPDX-License-Identifier: MIT
"""

import logging
import os

import pandas as pd
import win32com.client as win32
import config as cfg
from connections import Fuel
from connections import Grid
from files import clean_string
from files import read_input_parameter_data
from technologies import BHKW
from technologies import Boiler
from technologies import HeatPump
from technologies import Hybrid
from technologies import Source


class Scenario:
    def __init__(
        self,
        file,
        sheet_keyword="Inputparameter",
        r=None,
        base_year=None,
        period_under_consideration=None,
        flh=None,
        years=None,
    ):
        """
        A cost scenario for a group of technologies.

        Parameters
        ----------
        file : str
            Full path to a valid xlsx input file.
        sheet_keyword : str (optional)
            A substring to search for input data of different years within an
            input file.
        r : float (optional)
            Interest rate 'r'.
        base_year: int (optional)
            The base year for the net present value. If None the year of
            installation is used.
        period_under_consideration : int
            The period for the economic consideration.
        flh : int
            The full load hours for all technologies.
        years : list, array, Series
            The years to consider. Input data for all years must be present.
        """
        cfg.tmp_set("path", "fn", file)
        self.name = (
            os.path.basename(file).replace(".xlsx", "").replace(".xlsm", "")
        )
        self.technologies = {}
        self.economic_base = None
        self._results = None
        self.filename = os.path.basename(file)
        excel = win32.GetActiveObject("Excel.Application")
        wb = excel.Workbooks(self.filename)
        dashboard_datei = wb.Sheets("Dashboard")
        vls = dashboard_datei.Cells(5, 11).Value
        self.flh = [vls, vls+500]
        self.path = os.path.dirname(os.path.abspath(file))
        self.sheet_keyword = sheet_keyword
        self.input_parameter = read_input_parameter_data(file, sheet_keyword)
        if years is None:
            years = self.input_parameter.index.get_level_values(3).unique()

        self.years = years
        self._define_economic_base(r, base_year, period_under_consideration)
        self.add_technologies()

    def __str__(self):
        return "Scenario object: {0}".format(self.name)

    def _define_economic_base(self, r, base_year, period_under_consideration):
        """Create a dictionary for the basic economic values."""
        excel = win32.GetActiveObject("Excel.Application")
        wb = excel.Workbooks(self.filename)
        dashboard_datei = wb.Sheets("Dashboard")
        dc = {}
        if r is None:
            dc["r"] = dashboard_datei.Cells(6, 11).Value
        if base_year is None:
            dc["base year"] = dashboard_datei.Cells(7, 11).Value
        else:
            dc["base year"] = base_year
        if period_under_consideration is None:
            dc["period under consideration"] = cfg.get(
                "economic", "period under consideration"
            )
        self.economic_base = dc

    def create_result_table(self):
        """Calculate the results for all years and demand scenarios."""
        tables = []
        for year in self.years:
            msg = "Year: {0} with the following full load hours: {1}"
            logging.info(msg.format(year, self.flh))
            for flh in self.flh:
                for name, technology in self.technologies.items():
                    df = technology.heat_production_costs(year, flh)
                    for lev in [flh, year, name]:
                        df = pd.concat([df], axis=1, keys=[lev])
                    tables.append(df)
        self._results = pd.concat(tables, axis=1).T.stack().sort_index()
        if cfg.get("switch", "add_additional_result_objects") is True:
            self.add_wo_rows(
                "investment cost connection", " (Anschluss vorhanden)"
            )
            self.add_wo_rows(
                "investment cost heat source", " (Wärmequelle vorhanden)"
            )
        return tables


    def add_wo_rows(self, parameter, suffix):
        """Add rows with partly removed cost."""
        table_detailed = self._results.unstack([1, 2, 3, 4])
        parameter_cols = [c for c in table_detailed.columns if parameter in c]
        new_table = table_detailed.loc[
            table_detailed[parameter_cols].sum(axis=1) > 0
        ].copy()
        new_table.rename(
            {k: k + suffix for k in new_table.index},
            inplace=True,
        )
        new_table[parameter_cols] = 0
        self._results = (
            pd.concat([table_detailed, new_table])
            .stack([0, 1, 2, 3])
            .sort_index()
        )

    @property
    def results(self):
        if self._results is None:
            self.create_result_table()
        return self._results  # .loc[slice(5.0, 1500.0)]

    def results_to_xlsx(self, detailed=True, suffix=None, unit="EUR/kWh"):
        """
        Store results into a '.xlsx' file.

        Parameters
        ----------
        detailed : bool
            By default two files will be created with detailed and compact
            results. Setting detailed to False will only create one file with
            compact results.
        suffix : str
            A suffix to the resulting files to distinguish different scenarios.
        unit : str (default: "EUR/kWh")
            Valid units are: "EUR/kWh", "EUR/MWh", "ct/kWh".
        """
        path_detailed = os.path.join(
            self.path,
            "{0}_wgkpy_results_detailed.xlsx".format(clean_string(self.name)),
        )
        path = os.path.join(
            self.path,
            "{0}_wgkpy_results.xlsx".format(clean_string(self.name)),
        )

        if suffix is not None:
            path_detailed = path_detailed.replace(
                ".xlsx", "_{0}.xlsx".format(suffix)
            )
            path = path.replace(".xlsx", "_{0}.xlsx".format(suffix))

        table_detailed = (
            self.results.unstack([4, 3])
            .reorder_levels([1, 2, 0])
            .sort_index(axis=1)
            .sort_index(axis=0)
        )

        # Base unit EUR/kWh
        converter = {
            "EUR/kWh": 1,
            "EUR/MWh": 1000,
            "ct/kWh": 100,
        }

        if unit not in converter:
            msg = "Unknown unit: {0}\n Use one of the following units:{1}"
            raise ValueError(msg.format(unit, list(converter.keys())))

        # Write tables in EUR/MWh
        if detailed is True:
            table_detailed.mul(converter[unit]).to_excel(path_detailed)
            logging.info("Detailed results written to %s" % path_detailed)
        table = table_detailed.groupby(level=0, axis=1).sum()

        excel = win32.GetActiveObject("Excel.Application")
        wb = excel.Workbooks(str(self.name)+".xlsm")
        new_sheet = None
        sheet_name = "Ergebnisse"

        for sheet in wb.Sheets:
            if sheet.Name == sheet_name:
                new_sheet = sheet
                break

        if new_sheet is None:
            new_sheet = wb.Sheets.Add()
            new_sheet.Name = sheet_name

        ws = wb.Worksheets(sheet_name)
        StartRow = 1
        StartCol = 1
        table = table.reset_index().transpose().reset_index().transpose()
        ws.Range(ws.Cells(StartRow, StartCol), ws.Cells(StartRow + len(table.index) - 1,
                                                        StartCol + len(table.columns) - 1)).Value = table.values

        table.mul(converter[unit]).to_excel(path)
        logging.info("Compact results written to %s" % path)

    def create_objects(self, technologies, data):
        """Create wgkpy.Technology objects from input data."""
        if isinstance(technologies, dict):
            tech_types = assign_technologies(technologies.keys())
            hybrid_technology = list(technologies.values())[0]
            tech_types = {
                k: [{a: hybrid_technology} for a in v if len(v) > 0]
                for k, v in tech_types.items()
            }

        else:
            tech_types = assign_technologies(technologies)

        objects = []
        objects.extend(self._create_hybrid_objects(data, tech_types["hybrid"]))
        objects.extend(self._create_boiler_objects(data, tech_types["boiler"]))
        objects.extend(self._create_heat_pump_objects(data, tech_types["hp"]))
        objects.extend(self._create_source_objects(data, tech_types["source"]))
        objects.extend(self._create_chp_objects(data, tech_types["chp"]))
        return objects

    def add_technologies(self):
        """Add technologies to the Scenario class."""
        data = self.input_parameter
        data = data.droplevel(2)
        technologies_idx = data.index.get_level_values(0).unique()
        objects = self.create_objects(technologies_idx, data)
        self.technologies.update({o.name: o for o in objects})

    def _set_parameters(self, data, parameters, technology):
        parameter_names = cfg.get_dict("technology data")
        # Loop over the row names in the input table

        if isinstance(technology, dict):
            main = list(technology.values())[0]
            part = list(technology.keys())[0]
            for name in data[main].index.get_level_values(0).unique():
                if part in name:
                    keyname = name.replace(" " + part, "")
                    try:
                        parameters[parameter_names[keyname]] = data[main, name]
                    except KeyError as e:
                        msg = str(
                            "The parameter '{0}' is unknown. Fix input data "
                            "of {1}."
                        ).format(e.args[0], main)
                        raise KeyError(msg)

        else:
            for name in data[technology].index.get_level_values(0).unique():
                try:
                    parameters[parameter_names[name]] = data[technology, name]
                except KeyError as e:
                    msg = str(
                        "The parameter '{0}' is unknown. Fix input data of "
                        "{1}."
                    ).format(e.args[0], technology)
                    raise KeyError(msg)

        # Extract capacity classes from index
        parameters.setdefault(
            "capacity",
            pd.Series(
                parameters["investment_cost"].index.get_level_values(1),
                index=parameters["investment_cost"].index,
            ),
        )
        parameters["capacity"].loc[
            parameters["investment_cost"].isnull()
        ] = float("nan")
        if isinstance(technology, dict):
            name = list(technology.values())[0].split(" (")[0]
        else:
            name = technology
        parameters["name"] = name
        parameters["economic_base"] = self.economic_base
        parameters["fuel"] = Fuel(technology)
        return parameters

    def _create_boiler_objects(self, data, boilers):
        parameters = {}
        objects = []
        for boiler in boilers:
            parameters = self._set_parameters(data, parameters, boiler)
            objects.append(Boiler(**parameters))

        return objects

    def _create_source_objects(
        self,
        data,
        sources,
    ):

        parameters = {}
        objects = []
        for source in sources:
            self._set_parameters(data, parameters, source)
            parameters["fuel"] = None

            objects.append(Source(**parameters))
        return objects

    def _create_heat_pump_objects(
        self,
        data,
        heat_pumps,
    ):
        parameters = {}
        objects = []
        for heat_pump in heat_pumps:
            self._set_parameters(data, parameters, heat_pump)
            try:
                tmp_hp = HeatPump(**parameters)
            except TypeError as e:
                msg = "{0}\nFix input data of {1}".format(
                    e.args[0],
                    parameters.get("name"),
                )
                raise TypeError(msg)
            objects.append(tmp_hp)

        return objects

    def _create_chp_objects(
        self,
        data,
        bhkws,
    ):

        parameters = {}
        objects = []
        for bhkw in bhkws:
            self._set_parameters(data, parameters, bhkw)
            parameters["grid"] = Grid()
            tmp_chp = BHKW(**parameters)
            objects.append(tmp_chp)
        return objects

    def _create_hybrid_objects(self, data, hybrid_systems):

        objects = []
        for hybrid_tech in hybrid_systems:
            parts = hybrid_tech.split(" (")[-1].replace(")", "").split(",")
            parts = {p.strip(): hybrid_tech for p in parts}
            part_objs = self.create_objects(parts, data)

            sort_objs = {type(b): [] for b in part_objs}
            [sort_objs[type(b)].append(b) for b in part_objs]

            sort_objs_types = list(sort_objs.keys())

            for first in sort_objs[sort_objs_types[0]]:
                for second in sort_objs[sort_objs_types[1]]:
                    hybrid_parameters = {}
                    basename = first.name.split("_")[0]
                    if "_" in first.name:
                        first_extension = "_" + first.name.split("_")[1]
                    else:
                        first_extension = ""

                    if "_" in second.name:
                        second_extension = "_" + second.name.split("_")[1]
                    else:
                        second_extension = ""

                    name = basename + first_extension + second_extension
                    hybrid_parameters["technologies"] = [first, second]
                    hybrid_parameters["name"] = name
                    objects.append(Hybrid(**hybrid_parameters))

        return objects


def assign_technologies(technologie_names):
    """Assign input data sets to the wgkpy.Technology classes."""
    tech_types = {}

    remaining_names = list(technologie_names)

    # hybrid
    tech_types["hybrid"] = [
        h for h in remaining_names if "hybrid" in h.lower()
    ]
    remaining_names = [
        d for d in remaining_names if d not in tech_types["hybrid"]
    ]

    # heat pump
    tech_types["hp"] = [
        h for h in remaining_names if "wärmepumpe" in h.lower()
    ]
    remaining_names = [d for d in remaining_names if d not in tech_types["hp"]]

    # BHKW
    tech_types["chp"] = [h for h in remaining_names if "bhkw" in h.lower()]
    remaining_names = [
        d for d in remaining_names if d not in tech_types["chp"]
    ]

    # Source
    sources = ["Solar", "Abwärme", "Abw."]
    tmp = []
    for src in sources:
        tmp.extend([h for h in remaining_names if src in h])
    tech_types["source"] = tmp
    remaining_names = [
        d for d in remaining_names if d not in tech_types["source"]
    ]

    # boiler
    tech_types["boiler"] = [
        h
        for h in remaining_names
        if "kessel" in h.lower()
        or "fernw" in h.lower()
        or "infrarot" in h.lower()
    ]
    return tech_types
