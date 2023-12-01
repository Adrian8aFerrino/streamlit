# -*- coding: utf-8 -*-

"""
Incoming and outgoing connections.

SPDX-FileCopyrightText: Uwe Krien <uwe.krien@ifam.fraunhofer.de>

SPDX-License-Identifier: MIT
"""

import logging

import numpy as np
import pandas as pd

from files import read_fuel_matrix
from base import get_feedin_function
from base import get_price_function


class Fuel:
    """Fuel class."""

    def __init__(self, technology, name=None):
        if name is None:
            name = get_fuel_name(technology)
        self.name = name
        self.function = get_price_function(self.name)

    def get_price(self, years, demand):
        prices = {}
        for capacity in demand.index:
            if demand[capacity] > 0:
                prices[capacity] = self.function(years, demand[capacity])
        return pd.DataFrame(prices)


class Grid:
    """Electricity grid with sale and purchase cost functions."""

    def __init__(self):
        self.name = "electricity"
        self.sale_function = get_feedin_function(self.name)
        self.purchase_function = get_price_function(self.name)

    def sale(self, years):
        """Get sale prices for electricity for a given range of years."""
        prices = self.sale_function(years)
        return pd.Series(index=years, data=prices)

    def purchase(self, years, demand):
        """Get purchase prices for electricity for a given range of years."""
        prices = {}
        for capacity in demand.index:
            prices[capacity] = self.purchase_function(years, demand[capacity])
        return pd.DataFrame(prices)


def get_fuel_name(technology):
    """Assign the fuel name to a technology."""
    df = read_fuel_matrix()
    if isinstance(technology, dict):
        key = [(main.split(" (")[0], sub) for sub, main in technology.items()]
        fuel_name = df.loc[key[0], "Energieträger"]
        if isinstance(fuel_name, pd.Series):
            fuel_name = fuel_name.values[0]
    else:
        key = [(technology, "all")]
        fuel_name = df.loc[key[0], "Energieträger"]
        if isinstance(fuel_name, pd.Series):
            fuel_name = fuel_name.values[0]
    msg = "Fuel <{0}> assigned for {1}: {2}".format(
        fuel_name, key[0][0], key[0][1]
    )
    logging.debug(msg)
    return fuel_name
