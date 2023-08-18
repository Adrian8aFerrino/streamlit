import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Test Dashboard", page_icon=":frog:", layout="wide")
st.title("Test project of **Streamlit** Dashboard")


def logo_fraunhofer():
    st.markdown(
        """
            <style>
            [data-testid="stSidebarNav"] {
                background-image: url();
                background-repeat: no-repeat;
                padding-top: 120px;
                background-position: 20px 20px;
            }
            [data-testid="stSidebarNav"]::before {
                content: "Fraunhofer IFAM";
                margin-left: 20px;
                margin-top: 20px;
                font-size: 30px;
                position: relative;
                top: 100px;
            }
        </style>
        """)

uploaded_file = st.file_uploader(label="Upload a structured CSV file:", type=["csv"])

if uploaded_file is not None:
    data_test = pd.read_csv(uploaded_file, encoding="utf-8")
    categorical_variable = {}
    numerical_variable = {}
    other_variable = {}

    for column_name in data_test.columns:
        column_data = data_test[column_name]
        # Categorical variables
        if pd.api.types.is_categorical_dtype(column_data):
            categorical_variable[column_name] = column_data
        # Numerical variables
        elif pd.api.types.is_numeric_dtype(column_data):
            numerical_variable[column_name] = np.array(column_data)
        else:
            other_variable[column_name] = column_data

    print("Categorical Variables:")
    for column_name, column_data in categorical_variable.items():
        print(f"{column_name}:")
        print(column_data)
        print()

    # Print continuous variables
    print("Continuous Variables:")
    for column_name, column_data in numerical_variable.items():
        print(f"{column_name}:")
        print(column_data)
        print()

    # Print other variables
    print("Other Variables:")
    for column_name, column_data in other_variable.items():
        print(f"{column_name}:")
        print(column_data)
        print()

else:
    st.warning("Avoid the following Errors within your personal CSV files:")
    st.warning("ParserError: CSV file contains formatting issues. (missing delimiters)", icon="⚠️")
    st.warning("Encoding error: CSV file contains non-UTF-8 encoded characters.", icon="⚠️")
    st.warning("MemoryError: CSV file is too large to read.", icon="⚠️")
    st.warning("Among many others...")


