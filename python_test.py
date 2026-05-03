import config
import numpy as np
import pandas as pd
from datetime import datetime, timedelta


def dummy_energy_csv(n_rows):
    """
    Generates a csv file that can later be used within the streamlit application. The dataset contains variables
    related to an Energy Systems Analysis dataset, and is meant to be compatible with all three dashboard tabs. It
    contains the following variables:
    Kategoriale Daten:
    - Symbol
    - Company
    - Energy Source
    - Performance Rating
    - Region
    - Quarter

    Numerische Daten:
    - Commodity prices
    - Volume
    - CO2 Emissions
    - Grid Load MW
    - Renewable Share

    Zeitreihenanalyse:
    - Date
    - Crude Inventory Level
    """
    rng = np.random.default_rng(123)
    start_date = datetime(2024, 1, 1)
    dates = [start_date + timedelta(days=i) for i in range(n_rows)]
    months = np.array([d.month for d in dates])
    tag_jahr = np.array([d.timetuple().tm_yday for d in dates])
    weekdays = [d.weekday() for d in dates]
    seasons = ["Winter" if m in (12, 1, 2) else
               "Frühling" if m in (3, 4, 5) else
               "Sommer" if m in (6, 7, 8) else
               "Herbst"
               for m in months]

    # Categorical variables
    source_weights = {
        "Sommer": [0.55, 0.10, 0.05, 0.03, 0.15, 0.12],
        "Winter": [0.05, 0.20, 0.35, 0.28, 0.07, 0.05],
        "Frühling": [0.20, 0.40, 0.10, 0.05, 0.15, 0.10],
        "Herbst": [0.15, 0.45, 0.15, 0.08, 0.10, 0.07],
    }

    energy_sources = ["Solar", "Wind", "Gas", "Kohle", "Wasserkraft", "Biomasse"]
    sources = [rng.choice(energy_sources, p=source_weights[s]) for s in seasons]

    emission_factor = {"Solar": 5,
                       "Wind": 8,
                       "Wasserkraft": 4,
                       "Biomasse": 80,
                       "Gas": 400,
                       "Kohle": 820}

    price_base = {"Solar": 32,
                  "Wind": 28,
                  "Wasserkraft": 25,
                  "Biomasse": 55,
                  "Gas": 95,
                  "Kohle": 78}

    solar_base = 450 * np.clip(np.sin(2 * np.pi * (tag_jahr - 80) / 365), 0, None)
    solar_noise = rng.normal(0, 20, n_rows)
    solar = np.where(
        np.array(sources) == "Solar",
        solar_base * 1.4 + solar_noise,
        solar_base * 0.3 + solar_noise).clip(0)

    wind_base = 250 + 150 * np.sin(2 * np.pi * (tag_jahr - 265) / 365)
    wind_spikes = rng.exponential(60, n_rows) * (rng.random(n_rows) < 0.15)
    wind = np.where(
        np.array(sources) == "Wind",
        wind_base * 1.5 + wind_spikes + rng.normal(0, 30, n_rows),
        wind_base * 0.4 + rng.normal(0, 20, n_rows)).clip(0)

    seasonal_demand = 300 * np.sin(2 * np.pi * (tag_jahr - 15) / 365)
    weekend_dip = np.where(np.array(weekdays) >= 5, -80, 0)
    consumption = (900 + seasonal_demand + weekend_dip + rng.normal(0, 35, n_rows)).clip(200)
    emission_factor_arr = np.array([emission_factor[s] for s in sources], dtype=float)
    emission_co2 = ((emission_factor_arr * (consumption / 1000) * rng.uniform(0.85, 1.15, n_rows))
                    .clip(0).round(1))
    price_base_arr = np.array([price_base[s] for s in sources], dtype=float)
    winter_scarcity = np.where(np.isin(months, [12, 1, 2]), 25, 0)
    demand_pressure = (consumption - consumption.mean()) * 0.04
    price = (price_base_arr + winter_scarcity + demand_pressure + rng.normal(0, 5, n_rows)
             ).clip(5).round(2)
    p_overload = np.clip((emission_co2 / emission_co2.max()) * 0.5, 0.02, 0.45)
    p_underload = np.clip((1 - solar / (solar.max() + 1)) * 0.15, 0.01, 0.20)
    grid_status = []
    for i in range(n_rows):
        r = rng.random()
        if r < p_overload[i]:
            grid_status.append("Überlast")
        elif r < p_overload[i] + p_underload[i]:
            grid_status.append("Unterlast")
        else:
            grid_status.append("Stabil")
    p_maintenance = np.where(np.isin(months, [4, 5]), 0.22, 0.04)
    maintenance = np.where(rng.random(n_rows) < p_maintenance, "Ja", "Nein")

    region_weights_summer = [0.30, 0.15, 0.20, 0.20, 0.15]
    region_weights_winter = [0.15, 0.25, 0.25, 0.20, 0.15]
    regions = []
    for s in seasons:
        w = region_weights_summer if s == "Sommer" else region_weights_winter
        regions.append(
            rng.choice(["Nord", "Süd", "Ost", "West", "Mitte"], p=w))

    energy_dummy_df = pd.DataFrame({
        "Date": [d.strftime("%Y-%m-%d") for d in dates],
        "Jahreszeit": seasons,
        "Energiequelle": sources,
        "Region": regions,
        "Netzstatus": grid_status,
        "Wartung": maintenance,
        "Solarertrag_kWh": solar.round(2),
        "Windertrag_kWh": wind.round(2),
        "Verbrauch_kWh": consumption.round(2),
        "CO2_Emissionen_kg": emission_co2,
        "Preis_EUR_MWh": price,
    })

    float_cols = [
        "Solarertrag_kWh", "Windertrag_kWh",
        "Verbrauch_kWh", "CO2_Emissionen_kg", "Preis_EUR_MWh"]
    for col in float_cols:
        energy_dummy_df[col] = energy_dummy_df[col].astype(str).str.replace(".", ",", regex=False)

    return energy_dummy_df


destination_folder = config.select_folder("Databases")
energy_df = dummy_energy_csv(n_rows=730)
config.df2csv(destination_folder, energy_df, "xi_dataset")
