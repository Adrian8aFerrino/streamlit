import streamlit as st
import plotly.express as px
import pandas as pd
from PIL import Image

st.set_page_config(page_title="Test Dashboard", page_icon=":frog:", layout="wide")
st.title("Interaktives Dashboard des Wärmegestehungskostenmodell")

uploaded_file = st.file_uploader("Hier kannst du dein Excel Datein hochladen:")


def add_logo(logo_path, width, height):
    logo = Image.open(logo_path)
    modified_logo = logo.resize((width, height))
    return modified_logo


try:
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
    st.sidebar.header("Wikommen beim Fraunhofer IFAM")
    coil_filter = st.sidebar.multiselect("Select a category from dataframe_2 to filter:",
                                           options=list(data_test2["Condenser_Coil"].unique()),
                                           default=list(data_test2["Condenser_Coil"].unique()))
    symbol_filter = st.sidebar.multiselect("Select a category from dataframe_3 to filter:",
                                           options=list(data_test3["Symbol"].unique()),
                                           default=list(data_test3["Symbol"].unique()))

    tab1, tab2, tab3, tab4 = st.tabs(["Dataframes", "Plotting", "Analysis", "Python functions"])

    with tab1:
        text_uno = st.write("**Streamlit** is an open-source **:green[Python]** library that allows developers "
                            "to create interactive web applications for a wide variety of **:green[Python]** "
                            "related projects. **Streamlit** enables the creation of dashboards from "
                            "**:green[Python]** code. **Streamlit** can display excel tables as pandas "
                            "dataframes and apply **:green[Python]** code accordingly.")
        st.dataframe(data_test1)
        text_uno2 = st.write("**Streamlit** can select columns, filter data, and sort values using widgets like "
                             "dropdowns and sliders. **Streamlit** is reactive, meaning it automatically updates the "
                             "display whenever there is a change in user input. This feature is especially useful "
                             "when interacting with dataframes.")
        st.dataframe(data_test2)
        text_uno3 = st.write("**Streamlit** provides seamless integration with dataframes, allowing interaction with"
                             " them in a straightforward manner. It enables you to display specific columns or rows "
                             "based on user input. You can use widgets like checkboxes, select boxes, or radio "
                             "buttons to filter and display the desired data. ")
        st.dataframe(data_test3)
        text_uno4 = st.write("With **Streamlit**, you can easily display, manipulate, and visualize dataframes to "
                             "create interactive data-driven web applications. Whether it's for data exploration, "
                             "reporting, or data analysis, Streamlit offers a simple yet powerful solution for "
                             "working with dataframes in web applications.")
        st.dataframe(data_test4)

    with tab2:
        text_dos = st.write("**Streamlit** can also serve as a way to generate interactive graphs with the use of "
                            "common libraries like :red[Matplotlib], :red[Seaborn], and :red[Plotly]. These graphs "
                            "help communicate data effectively, in a data-driven application. Visualization options "
                            "like histograms, scatter plots, and line charts can be used to represent data "
                            "distribution and relationships.")
        dt3_selection = data_test3.query("Symbol == @symbol_filter")
        average_high = round(dt3_selection["High"].mean(), 2)
        average_low = round(dt3_selection["Low"].mean(), 2)
        average_volume = round(dt3_selection["Volume"].mean(), 2)
        average_performance = round(dt3_selection["Performance rating"].mean(), 0)

        left_col, mid_col, mid2_col, right_col = st.columns(4)
        with left_col:
            st.subheader("Average High rating")
            st.subheader(f"{average_high}")

        with mid_col:
            st.subheader("Average Low rating")
            st.subheader(f"{average_low}")

        with mid2_col:
            st.subheader("Average Volume rating")
            st.subheader(f"{average_volume} kg")

        with right_col:
            st.subheader("Average STAR rating")
            stars = ":star:" * round(int(average_performance),0)
            st.subheader(stars)

        open_by_company = dt3_selection.groupby("Company")["Open"].mean().reset_index()
        fig = px.bar(open_by_company, x="Company", y="Open", labels={"Open": "Average Open"})
        fig.update_traces(marker_color='orange')
        fig.update_layout(plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig)

        text_dos1 = st.write("**Streamlit** offers seamless interaction with plotting libraries, allowing you to "
                             "create dynamic and interactive visualizations in your web applications. You can also "
                             "leverage **Streamlit** widgets to modify the Plotly figure's attributes, update the "
                             "data, or switch between different types of charts.")

        data_test1['Datetime'] = pd.to_datetime(data_test1['Datetime'])
        fig = px.line(data_test1, x='Datetime', y=data_test1.columns[1:-1])
        st.plotly_chart(fig)

    with tab3:
        text_tres = st.write("**Streamlit** enables the user to display any type of :blue[KPI's] with the use of "
                             ":red[Pandas] and :red[NumPy] libraries as a way to manipulate and facilitates effective "
                             "data presentation.  Users can calculate summary statistics, run hypothesis tests, and "
                             "generate visualizations like box plots or violin plots to understand the distribution "
                             "of data.")
        brand_counts = data_test2['Brand_name'].value_counts().reset_index()
        brand_counts.columns = ['Brand_name', 'Count']
        fig = px.pie(brand_counts, values='Count', names='Brand_name', title='Interactive Pie Chart')
        st.plotly_chart(fig)
        text_four = st.write("Both :red[NumPy] and :red[Pandas] are powerful libraries in **:green[Python]** that are "
                             "widely used for advanced data manipulation, analysis, and processing. You can leverage "
                             "**Streamlit** widgets to modify the Plotly figure's attributes, update the data, or "
                             "switch between different types of charts.")
        st.dataframe(data_test5)
        if st.button("Generate Correlation map of the dataframe:"):
            correlation_matrix = data_test5.corr()
            fig = px.imshow(correlation_matrix,
                            x=correlation_matrix.columns,
                            y=correlation_matrix.index,
                            color_continuous_scale='Cividis',
                            title='Correlation Matrix')
            st.plotly_chart(fig)

    with tab4:
        text_cuatro = st.write("**Streamlit** allows the developer to display **:green[Python]** functions "
                               "interactively within a web applications, enabling users to see the outcomes of "
                               "functions in real-time. This can be particularly useful when demonstrating the "
                               "functionality of specific functions or libraries.")
        python_code = '''
        def python_code_example(env, client_number, client_data):
            while True:
                arrival_select = np.random.choice(new_arrival.values)
                client_data.loc[client_number, 'Arrival Rate'] = arrival_select
                if arrival_select != 0:
                    interarrival_time = 60 / arrival_select
                else:
                    interarrival_time = 60
                yield env.timeout(interarrival_time)
                client_number += 1
                station_mensa.put((env.now, client_number))
                env.process(service_process_mensa(env, client_data))
                '''
        st.code(python_code, language='python')

except:
    pass
