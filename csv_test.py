import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Test Dashboard", page_icon=":frog:", layout="wide")
st.title("Test project of **Streamlit** Dashboard")


import streamlit as st
import plotly.express as px
import pandas as pd
from PIL import Image
import numpy as np

st.set_page_config(page_title="Test Dashboard", page_icon="♻", layout="wide")
st.title("Interaktives Dashboard des **:green[Wärmegestehungskostenmodell]**")

uploaded_file = st.file_uploader("Hier kannst du dein Excel Datein hochladen:")


def add_logo(logo_path, width, height):
    logo = Image.open(logo_path)
    modified_logo = logo.resize((width, height))
    return modified_logo


if uploaded_file is not None:
    data_test1 = pd.read_excel(io=uploaded_file, engine="openpyxl", sheet_name="Demand_Watt", skiprows=3,
                               usecols=range(0, 8), nrows=8004)
    data_test2 = pd.read_excel(io=uploaded_file, engine="openpyxl", sheet_name="Air_Conditioner", skiprows=4,
                               usecols=range(8, 15), nrows=276)
    data_test3 = pd.read_excel(io=uploaded_file, engine="openpyxl", sheet_name="Oil_Gas", skiprows=1,
                               usecols=range(4, 13), nrows=23024)
    data_test4 = pd.read_excel(io=uploaded_file, engine="openpyxl", sheet_name="Demand_Watt", skiprows=1,
                               usecols=range(13, 14), nrows=34880)
    data_test5 = pd.read_excel(io=uploaded_file, engine="openpyxl", sheet_name="Correlation_Test", skiprows=0,
                               usecols=range(0, 14), nrows=60000)
else:
    st.warning("Upload a valid Excel file")

st.sidebar.image(add_logo(logo_path="logo.png", width=500, height=136))
st.sidebar.header("Willkommen beim Fraunhofer IFAM")
st.sidebar.write("Aktualisieren Sie DESTATIS-Daten automatisch aus dem Internet über die <request> library und "
                 "vorgegebene Anweisungen zur **:green[Datenextraktion]** im Statistischen Bundesamt.")
st.sidebar.divider()
st.sidebar.write("Generierung von **:green[Preiszeitreihen]** unter Verwendung von Preisannahmen, um eine dynamische "
                 "Darstellung der finanziellen Landschaft im Wärmeerzeugungssektor zu bieten, indem Schwankungen der "
                 "Gas- und Strompreis, Netzkosten, technologische Fortschritte und regulatorische Änderungen (Steuern, "
                 "Abgaben und Umlagen) erfasst werden.")
st.sidebar.divider()
st.sidebar.write("")

