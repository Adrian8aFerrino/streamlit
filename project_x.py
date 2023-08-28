import numpy as np
import pandas as pd
from PIL import Image
import streamlit as st
import statsmodels.api as sm
import matplotlib.pyplot as plt
from scipy.stats import skew, kurtosis

st.set_page_config(page_title="Streamlit Test Dashboard", page_icon=":frog:", layout="wide")
st.title("**:orange[Streamlit]** Test Dashboard")

uploaded_file = st.file_uploader(label="Upload a structured CSV file:", type=["csv"])


def add_logo(logo_path, width, height):
    logo = Image.open(logo_path)
    modified_logo = logo.resize((width, height))
    return modified_logo


def bar_chart(categorical_variable, title_var):
    bar_chart = plt.figure(figsize=(10, 6))
    frequency = categorical_variable.value_counts().sort_index()
    ax = frequency.plot(kind="bar", color=plt.cm.tab20c.colors)
    for p in ax.patches:
        ax.annotate(str(p.get_height()), (p.get_x() + p.get_width() / 2, p.get_height()), ha='center', va='bottom')
    ax.set_title(f"Bar plot for the frequency distribution of {title_var}")
    ax.set_ylabel("Frequency")
    ax.set_xlabel(f"- {title_var} -")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(bar_chart)


def pie_chart(categorical_variable, title_var):
    pie_chart = plt.figure(figsize=(4, 4))
    frequency = categorical_variable.value_counts().sort_index()
    plt.pie(frequency, labels=frequency.index, autopct='%1.1f%%', startangle=140, colors=plt.cm.tab20c.colors)
    plt.axis('equal')
    plt.title(f"Pie plot for the frequency distribution of {title_var}")
    st.pyplot(pie_chart)


def box_plot(numerical_variable, title_var):
    box_plot = plt.figure(figsize=(10, 6))
    ax = numerical_variable.plot(kind="box", color="blue")
    ax.set_title(f"Box plot for {title_var}")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(box_plot)


def time_series_plot(dataframe, title_var):
    dataframe['Date'] = pd.to_datetime(dataframe['Date'], format="%d.%m.%Y", errors='coerce')
    dataframe.dropna(subset=['Date'], inplace=True)
    dataframe.drop_duplicates(subset=['Date'], inplace=True)
    dataframe.set_index('Date', inplace=True)
    dataframe = dataframe.sort_index(ascending=True)
    data_test_num = dataframe[title_var].replace(",", ".", regex=True)
    data_test_num = data_test_num.astype(float)
    model = sm.tsa.arima.ARIMA(data_test_num, order=(1, 1, 1))
    results = model.fit()
    steps = len(dataframe.index) // 8
    forecast = results.forecast(steps=steps)
    time_series_plot = plt.figure(figsize=(10, 6))
    plt.plot(dataframe.index, data_test_num, label='Original Data', color="blue")
    plt.plot(pd.date_range(start=dataframe.index[-1], periods=steps, freq='D'), forecast, label='Forecast',
             color="orange")
    plt.title(f"Time Series plot for {title_var} with Forecast")
    plt.xlabel("Date")
    plt.ylabel(f"{title_var}")
    plt.legend()
    plt.tight_layout()
    st.pyplot(time_series_plot)


try:
    if uploaded_file is not None:
        data_test = pd.read_csv(uploaded_file, encoding="utf-8", sep=";")
        st.write(data_test)
        data_test_var = np.array(data_test.columns)
    
        tab1, tab2, tab3 = st.tabs(["Categorical data", "Numerical data", "Time series analysis"])
        with tab1:
            converted_columns = []
            for column in data_test.columns:
                num_categories = data_test[column].nunique()
                if num_categories <= 12:
                    converted_columns.append(column)
    
            categorical_data = st.selectbox("Select categorical data that you would want to be analyzed:",
                                            converted_columns)
            data_test_cat = data_test[categorical_data].astype("category")
            st.write("Exploring categorical data involves analyzing and summarizing data that falls into distinct "
                     "categories or groups. A frequency distribution is highly useful when dealing with categorical data."
                     " It provides a clear and concise summary of how the different categories or values within a "
                     "categorical variable are distributed within a dataset.")
    
            bar_button = st.button('Generate Bar plot')
            if bar_button:
                st.write(bar_chart(data_test_cat, categorical_data))
    
            pie_button = st.button('Generate Pie plot')
            if pie_button:
                st.write(pie_chart(data_test_cat, categorical_data))
    
            st.write("When not all categories are equally represented in a categorical variable, it can have several "
                     "negative implications for data analysis and interpretation. This is due to the fact that an "
                     "**:blue[Skewed discernment]** of the overall data patterns happens when a category dominates the "
                     "distribution, generating **:blue[Bias]** in the analysis and a faulty "
                     "**:blue[Statistical significance]**, which in turn impacts **:blue[Model performance]**.")
    
        with tab2:
            converted_columns = []
            for col in data_test.columns:
                try:
                    data_test[col] = data_test[col].replace(",", ".", regex=True)
                    data_test[col] = data_test[col].astype(float)
                    converted_columns.append(col)
                except ValueError:
                    pass
    
            numerical_data = st.selectbox("Select numerical data that you would want to be analyzed:", converted_columns)
            data_test_num = data_test[numerical_data].replace(",", ".", regex=True)
            data_test_num = data_test_num.astype(float)
            st.write("Exploring numerical data involves analyzing and summarizing data skewness, central tendency, and "
                     "variability of a variable. Summary statistics are immensely useful in data analysis for providing a "
                     "concise and insightful overview of a dataset's characteristics. They help in understanding the "
                     "central tendency, variability, and distribution of numerical data.")
            mean = data_test_num.mean()
            median = data_test_num.median()
            mode = data_test_num.mode().iloc[0]
            std_deviation = data_test_num.std()
            variance = data_test_num.var()
            quantiles = data_test_num.quantile([0.25, 0.5, 0.75])
            skewness = skew(data_test_num)
            kurtosis = kurtosis(data_test_num)
    
            col1, col2, col3, col4 = st.columns(4)
    
            with col1:
                st.write("**Mean**")
                st.write("**Median**")
                st.write("**Mode**")
                st.write("**Variance**")
                st.write("**Standard Deviation**")
    
            with col2:
                st.write(mean)
                st.write(median)
                st.write(mode)
                st.write(std_deviation)
                st.write(variance)
    
            with col3:
                st.write("**Quantile 25%**")
                st.write("**Quantile 50%**")
                st.write("**Quantile 75%**")
                st.write("**Skewness**")
                st.write("**Kurtosis**")
    
            with col4:
                st.write(quantiles.loc[0.25])
                st.write(quantiles.loc[0.5])
                st.write(quantiles.loc[0.75])
                st.write(skewness)
                st.write(kurtosis)
    
            box_button = st.button('Generate Box plot')
            if box_button:
                st.write(box_plot(data_test_num, numerical_data))
    
            st.write("A box plot provides a clear and concise summary of the distribution and variability of a dataset and "
                     "are particularly useful for visualizing and comparing the distribution of numerical data. Within a "
                     "box plot one can determine the skewness and symmetry of the distribution by comparing the length of "
                     "the whiskers and the position of the median and quartiles")
            st.write("Since many statistical methods assume normal distribution. Skewed data might require appropriate "
                     "transformations in order to be mathematically tractable and ensure the validity of inferential "
                     "statistical tests like the **:blue[ANOVA]** statistical test, **:blue[T-tests]** or "
                     "**:blue[Correlation analysis]**.")
    
        with tab3:
            converted_columns = []
            for col in data_test.columns:
                try:
                    data_test[col] = data_test[col].replace(",", ".", regex=True)
                    data_test[col] = data_test[col].astype(float)
                    converted_columns.append(col)
                except ValueError:
                    pass
    
            numerical_data_2 = st.selectbox("Select data for Time Series plot:", converted_columns)
            data_test_num = data_test[numerical_data_2].replace(",", ".", regex=True)
            data_test_num = data_test_num.astype(float)
    
            st.write("Time series analysis revolves around the examination of historical data points over time "
                     "intervals by examining **:blue[seasonality]**, **:blue[trends]** and **:blue[fluctuations]** "
                     "within the data, that can be later examined with techniques such as ARIMA models.")
            st.write("For this interactive dashboard we will employ an ARIMA model which stands for "
                     "**:blue[Autoregressive (AR)]**, **:blue[Integrated (I)]** and **:blue[Moving Average (MA)]** "
                     "that are particularly effective for capturing complex patterns, trends, and seasonality in time "
                     "series data. **:blue[(The model contains fixed p,d,q values)]**")

            if 'Date' in data_test.columns:
                time_series_button = st.button('Generate Time Series plot')
                if time_series_button:
                    st.write(time_series_plot(data_test, numerical_data_2))
                    st.code('''def time_series_plot(dataframe, title_var):
                    dataframe['Date'] = pd.to_datetime(dataframe['Date'], format="%d.%m.%Y", errors='coerce')
                    dataframe.dropna(subset=['Date'], inplace=True)
                    dataframe.drop_duplicates(subset=['Date'], inplace=True)
                    dataframe.set_index('Date', inplace=True)
                    dataframe = dataframe.sort_index(ascending=True)
                    data_test_num = dataframe[title_var].replace(",", ".", regex=True)
                    data_test_num = data_test_num.astype(float)
                    model = sm.tsa.arima.ARIMA(data_test_num, order=(1, 1, 1))
                    results = model.fit()
                    steps = len(dataframe.index) // 8
                    forecast = results.forecast(steps=steps)
                    time_series_plot = plt.figure(figsize=(10, 6))
                    plt.plot(dataframe.index, data_test_num, label='Original Data', color="blue")
                    plt.plot(pd.date_range(start=dataframe.index[-1], periods=steps, freq='D'), forecast, label='Forecast',
                             color="orange")
                    plt.title(f"Time Series plot for {title_var} with Forecast")
                    plt.xlabel("Date")
                    plt.ylabel(f"{title_var}")
                    plt.legend()
                    plt.tight_layout()
                    st.pyplot(time_series_plot)''', language="python")
            else:
                time_series_button = st.button('No -Date- column within CSV file')
    
    else:
        st.warning("Avoid the following Errors within your personal CSV files:")
        st.warning("ParserError: CSV file contains formatting issues. (missing delimiters)", icon="⚠️")
        st.warning("Encoding error: CSV file contains non-UTF-8 encoded characters.", icon="⚠️")
        st.warning("MemoryError: CSV file is too large to read.", icon="⚠️")
        st.warning("Among many others...")
    
    st.sidebar.image(add_logo(logo_path="project_xi.png", width=500, height=500))
    st.sidebar.header("Welcome to the **:orange[Streamlit]** Dashboard for data analytics.")
    st.sidebar.write("Within this dashboard you can experiment with several fundamental measurements, tests, and "
                     "graphs that are often shown in a data report.")
    st.sidebar.divider()
    st.sidebar.write("Business intelligence and data exploration are essential for obtaining meaningful insights from "
                     "data, whether it is for recognizing trends and effectively conveying findings.")
    st.sidebar.divider()
    st.sidebar.write("An interactive dashboard bridges the gap between theoretical and abstract methodologies "
                     "employed in data analytics and the concrete manifestations of these methods when applied to "
                     "real-world data.")
    
except:
    pass
