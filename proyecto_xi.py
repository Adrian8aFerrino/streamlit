import warnings
import numpy as np
import pandas as pd
from PIL import Image
import seaborn as sns
import streamlit as st
import matplotlib.pyplot as plt
from scipy.stats import skew, kurtosis
from statsmodels.tsa.stattools import adfuller
from statsmodels.tools.sm_exceptions import ConvergenceWarning
from statsmodels.tsa.seasonal import STL
from statsmodels.tsa.holtwinters import ExponentialSmoothing

warnings.simplefilter('ignore', ConvergenceWarning)

st.set_page_config(page_title="Streamlit Test Dashboard", page_icon=":frog:", layout="wide")
st.title("**:orange[Streamlit]** Test Dashboard")

uploaded_file = st.file_uploader(label="Lade eine strukturierte CSV-Datei hoch:", type=["csv"])


def add_logo(logo_path, width, height):
    logo = Image.open(logo_path)
    modified_logo = logo.resize((width, height))
    return modified_logo


def bar_chart(categorical_variable, title_var):
    fig, ax = plt.subplots(figsize=(10, 6))
    frequency = categorical_variable.value_counts().sort_index()
    bars = ax.bar(frequency.index, frequency.values, color=sns.color_palette("tab20c"))

    ax.set_title(f"Häufigkeitsverteilung: {title_var}")
    ax.set_ylabel("Anzahl")
    ax.set_xlabel(title_var)
    ax.tick_params(axis='x', rotation=45)

    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{int(height)}', xy=(bar.get_x() + bar.get_width() / 2, height),
                    ha='center', va='bottom')

    st.pyplot(fig)


def pie_chart(categorical_variable, title_var):
    frequency = categorical_variable.value_counts().sort_index()
    fig, ax = plt.subplots(figsize=(5, 5))
    wedges, texts, autotexts = ax.pie(
        frequency, labels=frequency.index, autopct='%1.1f%%', startangle=90,
        colors=sns.color_palette("pastel"), textprops={'fontsize': 10}
    )
    ax.set_title(f"{title_var}: Anteil nach Kategorie")
    st.pyplot(fig)


def box_plot(numerical_variable, title_var):
    fig, ax = plt.subplots(figsize=(8, 6))
    sns.boxplot(y=numerical_variable, color="lightblue", ax=ax)
    ax.set_title(f"Boxplot: {title_var}")
    ax.set_ylabel(title_var)
    st.pyplot(fig)


def histogram_plot(numerical_variable, title_var, bins=30):
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.histplot(numerical_variable, kde=True, bins=bins, ax=ax, color="skyblue", edgecolor="black")
    ax.set_title(f"Histogramm: {title_var}")
    ax.set_xlabel(title_var)
    ax.set_ylabel("Häufigkeit")
    st.pyplot(fig)


def time_series_plot(dataframe, column):
    # ── 1. Clean & prepare ────────────────────────────────────────────────────
    dataframe = dataframe.copy()
    dataframe['Date'] = pd.to_datetime(dataframe['Date'], errors='coerce')
    dataframe = (dataframe
                 .dropna(subset=['Date'])
                 .drop_duplicates('Date')
                 .sort_values('Date'))
    dataframe.set_index('Date', inplace=True)

    # Parse European decimal comma → float
    series = (dataframe[column]
              .astype(str)
              .str.replace(',', '.', regex=False))
    series = pd.to_numeric(series, errors='coerce').dropna()

    # ── 2. Enforce daily frequency & interpolate any gaps ────────────────────
    series = series.asfreq('D')
    series = series.interpolate(method='time', limit_direction='forward').dropna()

    if len(series) < 30:
        st.warning("Nicht genügend Datenpunkte für eine zuverlässige Vorhersage.")
        return

    # ── 3. Training window: last 90 days; plot window: last 180 days ─────────
    cutoff_model = series.index.max() - pd.Timedelta(days=90)
    cutoff_plot = series.index.max() - pd.Timedelta(days=180)
    model_series = series.loc[series.index >= cutoff_model]
    plot_series = series.loc[series.index >= cutoff_plot]

    # ── 4. ADF stationarity check ─────────────────────────────────────────────
    adf_p = adfuller(model_series)[1]
    st.write(f"**ADF-Test p-Wert:** {adf_p:.4f}")
    if adf_p > 0.05:
        st.warning("Nicht stationär – Trend/Saisonkomponenten werden durch "
                   "Holt-Winters modelliert.")
    else:
        st.success("Stationär – Holt-Winters kann direkt angewendet werden.")

    # ── 5. STL decomposition – diagnostic view ───────────────────────────────
    with st.expander("STL-Zerlegung anzeigen (Trend · Saisonalität · Residuum)"):
        try:
            stl = STL(model_series, period=7, robust=True).fit()
            fig_stl, axes = plt.subplots(4, 1, figsize=(12, 8), sharex=True)
            axes[0].plot(model_series, color='steelblue')
            axes[0].set_ylabel('Original')
            axes[1].plot(stl.trend, color='darkorange')
            axes[1].set_ylabel('Trend')
            axes[2].plot(stl.seasonal, color='seagreen')
            axes[2].set_ylabel('Saisonalität (7 Tage)')
            axes[3].plot(stl.resid, color='grey')
            axes[3].set_ylabel('Residuum')
            fig_stl.suptitle(f"STL-Zerlegung: {column}", fontsize=13)
            plt.tight_layout()
            st.pyplot(fig_stl)
        except Exception as e:
            st.info(f"STL-Zerlegung nicht verfügbar: {e}")

    # ── 6. Holt-Winters forecast ──────────────────────────────────────────────
    try:
        with st.spinner("Holt-Winters-Modell wird angepasst …"):
            hw_model = ExponentialSmoothing(
                model_series,
                trend='add',
                seasonal='add',
                seasonal_periods=7,
                damped_trend=True,
                initialization_method='estimated',
            ).fit(optimized=True)

        steps = 30
        forecast = hw_model.forecast(steps)
        forecast_dates = pd.date_range(
            start=model_series.index[-1] + pd.Timedelta(days=1),
            periods=steps, freq='D'
        )

        # 95 % confidence interval via in-sample residual std
        resid_std = np.std(hw_model.resid)
        ci_lower = forecast - 1.96 * resid_std
        ci_upper = forecast + 1.96 * resid_std

        # ── 7. Plot ───────────────────────────────────────────────────────────
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(plot_series.index, plot_series,
                label='Historisch', color='steelblue', linewidth=1.5)
        ax.axvline(model_series.index[0], color='grey', linestyle='--',
                   linewidth=0.8, label='Trainingsstart (−90 Tage)')
        ax.plot(forecast_dates, forecast,
                label='Prognose (30 Tage)', color='darkorange', linewidth=2)
        ax.fill_between(forecast_dates, ci_lower, ci_upper,
                        color='orange', alpha=0.25,
                        label='95 % Konfidenzintervall')
        ax.set_title(f"Holt-Winters Prognose für {column} "
                     f"(saisonale Periode = 7 Tage)", fontsize=13)
        ax.set_xlabel("Datum")
        ax.set_ylabel(column)
        ax.legend()
        plt.tight_layout()
        st.pyplot(fig)

        mae = np.mean(np.abs(hw_model.resid))
        rmse = np.sqrt(np.mean(hw_model.resid ** 2))
        col1, col2 = st.columns(2)
        col1.metric("MAE  (In-Sample)", f"{mae:.2f}")
        col2.metric("RMSE (In-Sample)", f"{rmse:.2f}")

    except Exception as e:
        st.error(f"Modellfehler: {e}")


try:
    if uploaded_file is not None:
        data_test = pd.read_csv(uploaded_file, encoding="utf-8", sep=";")
        data_test = data_test.loc[:, ~data_test.columns.str.match(r'^Unnamed')]
        st.write(data_test)
        data_test_var = np.array(data_test.columns)

        tab1, tab2, tab3 = st.tabs(["Kategoriale Daten", "Numerische Daten", "Zeitreihenanalyse"])

        with tab1:
            converted_columns = []
            for column in data_test.columns:
                num_categories = data_test[column].nunique()
                if num_categories <= 12:
                    converted_columns.append(column)

            categorical_data = st.selectbox("Die kategorialen Daten auswählen:",
                                            converted_columns)
            data_test_cat = data_test[categorical_data].astype("category")
            st.write("Die Untersuchung kategorialer Daten umfasst die Analyse und Zusammenfassung von Daten, "
                     "die in verschiedene Kategorien oder Gruppen fallen. Eine Häufigkeitsverteilung ist bei "
                     "der Arbeit mit kategorialen Daten sehr nützlich. Sie bietet eine klare und prägnante "
                     "Zusammenfassung darüber, wie die verschiedenen Kategorien oder Werte innerhalb einer "
                     "kategorialen Variablen innerhalb eines Datensatzes verteilt sind.")

            bar_button = st.button('Balkendiagramm erstellen')
            if bar_button:
                bar_chart(data_test_cat, categorical_data)

            st.write("Das Dashboard enthält zwei wichtige visuelle Tools zur Analyse kategorialer Variablen: "
                     "Balkendiagramme und Kreisdiagramme. Balkendiagramme bieten eine übersichtliche Darstellung "
                     "der Häufigkeitsverteilung von Kategorien und eignen sich daher ideal zum Erkennen "
                     "dominanter Gruppen, seltener Ereignisse und Ungleichgewichte in den Daten. Diese Diagramme "
                     "unterstützen die vergleichende Analyse zwischen Gruppen und sind besonders nützlich bei "
                     "der Interpretation nominaler oder ordinaler Variablen. Im Gegensatz dazu betonen "
                     "Kreisdiagramme die relativen Anteile jeder Kategorie und bieten eine intuitivere "
                     "Darstellung, wenn es darum geht, zu vermitteln, wie viel jede Gruppe zum Ganzen beiträgt. "
                     "Zusammen helfen diese Visualisierungen den Benutzern, die Struktur, Ausgewogenheit und "
                     "potenzielle Verzerrung kategorialer Merkmale zu bewerten, was bei der Vorbereitung von "
                     "Daten für die Modellierung oder Interpretation von entscheidender Bedeutung ist.")

            pie_button = st.button('Kreisdiagramm erstellen')
            if pie_button:
                pie_chart(data_test_cat, categorical_data)

            st.write("Wenn nicht alle Kategorien in einer kategorialen Variablen gleichermaßen vertreten sind, "
                     "kann dies mehrere negative Auswirkungen auf die Datenanalyse und -interpretation haben. Dies "
                     "liegt daran, dass es zu einer **:blue[verzerrten Wahrnehmung]** der Gesamtdatenmuster kommt, "
                     "wenn eine Kategorie die Verteilung dominiert, was zu einer Verzerrung der Analyse und einer "
                     "fehlerhaften **:blue[statistischen Signifikanz]** führt, was sich wiederum auf die "
                     "**:blue[Modellleistung auswirkt]**.")

        with tab2:
            # Work on a copy to avoid mutating data_test for Tab 3
            data_test_copy = data_test.copy()
            converted_columns = []
            for col in data_test_copy.columns:
                try:
                    data_test_copy[col] = (data_test_copy[col]
                                           .astype(str)
                                           .str.replace(",", ".", regex=False)
                                           .astype(float))
                    converted_columns.append(col)
                except ValueError:
                    pass

            numerical_data = st.selectbox("Die numerische Daten auswählen:",
                                          converted_columns)
            data_test_num = data_test_copy[numerical_data]

            st.write("Das Dashboard bietet mehrere Visualisierungsmethoden zur Untersuchung numerischer Daten: "
                     "Boxplots, Histogramme und Zeitreihendiagramme. Boxplots bieten eine kompakte Zusammenfassung "
                     "der Verteilungsmerkmale – wie Median, Streuung, Schiefe und Ausreißer –, sodass Benutzer "
                     "Anomalien erkennen und die Variabilität auf einen Blick beurteilen können. Histogramme "
                     "ergänzen Boxplots, indem sie die Häufigkeit von Werten über definierte Intervalle (Bins) "
                     "hinweg veranschaulichen, was Benutzern hilft, die Form, Modalität und Schiefe der Verteilung "
                     "zu verstehen. Schließlich kombiniert das Zeitreihendiagramm Datenvisualisierung mit Prognosen, "
                     "indem es ein Holt-Winters-Modell (Dreifache Exponentielle Glättung) verwendet, um Trends und "
                     "wöchentliche Saisonalität zu erfassen und zukünftige Werte zu prognostizieren, wobei "
                     "Unsicherheiten durch Konfidenzintervalle berücksichtigt werden. "
                     "Diese Diagramme bilden ein umfassendes Toolkit für die Analyse numerischer Variablen, die "
                     "Validierung von Annahmen und die Generierung prädiktiver Erkenntnisse.")

            mean = data_test_num.mean()
            median = data_test_num.median()
            mode = data_test_num.mode().iloc[0]
            std_deviation = data_test_num.std()
            variance = data_test_num.var()
            quantiles = data_test_num.quantile([0.25, 0.5, 0.75])
            skewness = skew(data_test_num)
            kurt_val = kurtosis(data_test_num)

            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.write("**Mittelwert**")
                st.write("**Median**")
                st.write("**Modus**")
                st.write("**Varianz**")
                st.write("**Standardabweichung**")

            with col2:
                st.write(mean)
                st.write(median)
                st.write(mode)
                st.write(std_deviation)
                st.write(variance)

            with col3:
                st.write("**Quantil 25%**")
                st.write("**Quantil 50%**")
                st.write("**Quantil 75%**")
                st.write("**Schiefe**")
                st.write("**Kurtosis**")

            with col4:
                st.write(quantiles.loc[0.25])
                st.write(quantiles.loc[0.5])
                st.write(quantiles.loc[0.75])
                st.write(skewness)
                st.write(kurt_val)  # fixed: was incorrectly referencing the kurtosis function object

            box_button = st.button('Box plot erstellen')
            if box_button:
                box_plot(data_test_num, numerical_data)

            st.write("Die Boxplot-Funktion ist ein zentrales Werkzeug bei der statistischen Datenauswertung. Sie "
                     "fasst die Verteilung einer numerischen Variablen anhand wichtiger Quantile (Minimum, Q1, "
                     "Median, Q3, Maximum) zusammen und hebt potenzielle Ausreißer hervor. Boxplots sind "
                     "entscheidend für die schnelle Beurteilung der Symmetrie, Schiefe und Streuung von Daten "
                     "sowie für die Identifizierung von Extremwerten, die sich auf nachgelagerte Modelle oder "
                     "Analysen auswirken könnten. Sie sind besonders wertvoll, wenn es darum geht, die Verteilungen"
                     " mehrerer Variablen zu vergleichen oder die Konsistenz von Messungen über einen bestimmten "
                     "Zeitraum oder über verschiedene Kategorien hinweg zu verstehen.")

            histogram_button = st.button('Histogramm erstellen')
            if histogram_button:
                histogram_plot(data_test_num, numerical_data)

            st.write("Wichtig: Viele statistische Methoden gehen davon aus, dass die Daten normalverteilt sind. "
                     "Wenn Daten nicht normalverteilt sind, können Transformationen erforderlich sein, um "
                     "diese Annahmen zu erfüllen und die Gültigkeit von Tests wie **:blue[ANOVA]**, "
                     "**:blue[t-tests]** und **:blue[Korrelationsanalysen]** sicherzustellen.")

        with tab3:
            # Detect numeric columns: try-convert to handle comma-decimal CSVs
            data_test_tab3 = data_test.copy()
            converted_columns = []
            for col in data_test_tab3.columns:
                if col == 'Date':
                    continue
                if pd.api.types.is_numeric_dtype(data_test_tab3[col]):
                    converted_columns.append(col)
                else:
                    try:
                        data_test_tab3[col] = (data_test_tab3[col]
                                               .astype(str)
                                               .str.replace(',', '.', regex=False)
                                               .astype(float))
                        converted_columns.append(col)
                    except (ValueError, AttributeError):
                        pass

            numerical_data_2 = st.selectbox("Daten für Zeitreihendiagramm auswählen:", converted_columns)

            st.write("Die Zeitreihenanalyse befasst sich mit der Untersuchung historischer Datenpunkte über "
                     "Zeitintervalle hinweg, indem die **:blue[Saisonalität]**, **:blue[Trends]** und "
                     "**:blue[Schwankungen]** innerhalb der Daten untersucht werden, die mit dem "
                     "Holt-Winters-Modell (Dreifache Exponentielle Glättung) modelliert werden können.")
            st.write("Für dieses interaktive Dashboard verwenden wir ein **:blue[Holt-Winters-Modell]**, das für "
                     "**:blue[Niveau (Level)]**, **:blue[Trend]** und **:blue[Saisonalität]** steht. Es ist "
                     "besonders effektiv, um wöchentliche Muster und Trendverschiebungen in Zeitreihendaten zu "
                     "erfassen. **:blue[(Die saisonale Periode ist auf 7 Tage festgelegt)]**")
            st.write("Mit der Zeitreihen-Plot-Funktion können Benutzer visualisieren, wie sich eine numerische "
                     "Variable im Laufe der Zeit entwickelt, und diese Analyse mithilfe des Holt-Winters-Modells "
                     "auf die Zukunft ausweiten. Zusätzlich bietet die STL-Zerlegung (Seasonal-Trend "
                     "Decomposition) eine diagnostische Ansicht, die Trend, wöchentliche Saisonalität und "
                     "Residuum separat darstellt. Die Zeitreihenanalyse ist unerlässlich, um zeitliche Trends, "
                     "Saisonalität, zyklisches Verhalten und Musterverschiebungen aufzudecken.")

            if 'Date' in data_test.columns:
                time_series_button = st.button('Zeitreihen-Diagramm erstellen')
                if time_series_button:
                    time_series_plot(data_test_tab3, numerical_data_2)
            else:
                st.button('Keine „Date"-Spalte in der CSV-Datei', disabled=True)

    else:
        st.warning("Vermeiden Sie die folgenden Fehler in Ihren persönlichen CSV-Dateien:")
        st.warning("ParserError: CSV-Datei enthält Formatierungsfehler. (fehlende Trennzeichen)", icon="⚠️")
        st.warning("Kodierungsfehler: CSV-Datei enthält Zeichen, die nicht in UTF-8 kodiert sind.", icon="⚠️")
        st.warning("Speicherfehler: Die CSV-Datei ist zu groß, um gelesen zu werden.", icon="⚠️")
        st.warning("Unter vielen anderen...")

    st.sidebar.image(add_logo(logo_path="project_xi.png", width=500, height=500))
    st.sidebar.header("Willkommen beim **:orange[Streamlit]** Dashboard für Datenanalysen.")
    st.sidebar.write("Das Dashboard dient dem Experimentieren mit verschiedenen grundlegenden Messungen, Tests und "
                     "Grafiken, die häufig in Datenberichten zur Anzeige kommen.")
    st.sidebar.divider()
    st.sidebar.write("Business Intelligence und Datenauswertung sind unerlässlich, um aussagekräftige Erkenntnisse "
                     "aus Daten zu gewinnen - sei es zur Erkennung von Trends oder zur effektiven Vermittlung von "
                     "Ergebnissen.")
    st.sidebar.divider()
    st.sidebar.write("Ein interaktives Dashboard schließt die Lücke zwischen den theoretischen und abstrakten "
                     "Methoden der Datenanalyse und den konkreten Ergebnissen, die sich ergeben, wenn diese "
                     "Methoden auf reale Daten angewendet werden.")

except Exception as e:
    st.error(f"Unerwarteter Fehler: {e}")