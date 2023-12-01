# -*- coding: utf-8 -*-

"""
Access to all price functions.

SPDX-FileCopyrightText: Esmail Ansari <esmail.ansari@ifam.fraunhofer.de
SPDX-FileCopyrightText: Uwe Krien <uwe.krien@ifam.fraunhofer.de>

SPDX-License-Identifier: MIT
"""
import pandas as pd

import config as cfg
from district_heating import get_district_heating_prices
from electricity import get_power_feedin_prices
from electricity import get_power_prices
from electricity import get_wp_power_prices
from gas import get_biogas_prices
from gas import get_gas_prices
from gas import get_h2_prices
from gas import get_mixgas_prices
from other import get_oil_prices
from other import get_pellet_prices


def get_price_function(fuel):
    translate = cfg.get_dict("fuels")
    if fuel in translate.keys():
        fuel = translate[fuel]

    if fuel is None or fuel == "None":
        return None

    price_functions = {
        "natural gas": get_gas_prices,
        "biogas": get_biogas_prices,
        "h2": get_h2_prices,
        "electricity": get_power_prices,
        "mixed gas": get_mixgas_prices,
        "wood pellets": get_pellet_prices,
        "oil": get_oil_prices,
        "electricity hp": get_wp_power_prices,
        "district heating": get_district_heating_prices,
    }

    if fuel not in price_functions.keys():
        print(f"{fuel} is not included in the price function")
        fuel = "mixed gas"

    if fuel not in price_functions.keys():
        msg = (
            "There is no price data available for fuel '{0}'.\nUse one of the "
            "following fuel types:\n{1}\n{2}"
        )
        raise ValueError(
            msg.format(
                fuel, list(price_functions.keys()), list(translate.keys())
            )
        )
    return price_functions[fuel]


def get_feedin_function(medium):
    price_functions = {
        "electricity": get_power_feedin_prices,
    }
    if medium not in price_functions.keys():
        msg = (
            "There is no feed-in data available for fuel '{0}'.\nUse one of "
            "the following feed-in types: {1}"
        )
        raise ValueError(msg.format(medium, list(price_functions.keys())))
    return price_functions[medium]


def get_feedin_prices(fuel, years):
    """
    Get feed-in prices for a specific fuel and a given range of years.

    Parameters
    ----------
    fuel : str
        Name of the fuel.
    years : list, array
        The range of years to get prices for.

    Returns
    -------
    pandas.Series
        One price for each year.
    """
    price_functions = {
        "electricity": get_power_feedin_prices,
    }
    if fuel not in price_functions.keys():
        msg = (
            "There is no feed-in data available for fuel '{0}'.\nUse one of "
            "the following feed-in types: {1}"
        )
        raise ValueError(msg.format(fuel, list(price_functions.keys())))
    prices = price_functions[fuel](years)

    return pd.Series(index=years, data=prices)


def get_fuel_prices(fuel, years, demand):
    """
    Get feed-in prices for a specific fuel and a given range of years.

    Parameters
    ----------
    fuel : str
        Name of the fuel e.g. natural gas, oil, electricity.
    years : list, array
        The range of years to get prices for.
    demand : numeric
        The annual demand of the fuel.

    Returns
    -------
    pandas.Series
        One price for each year.
    """
    price_function = get_price_function(fuel)
    prices = price_function(years, demand)

    return prices
