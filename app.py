import io
import random
from io import BytesIO
import datetime
from datetime import timedelta
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
from PIL.ImageChops import overlay
from dateutil.utils import today

import streamlit as st
from helper_functions import style_dataframe, load_profitbility_Summary_model, load_rates_standardisation, \
    load_specific_xls_sheet, sub_category_classification, style_commodity_customers
from streamlit import session_state as ss
from plotly.subplots import make_subplots

from site_detailed_view_helper_functions import load_invoices_model, profitability_model

st.set_page_config(page_title="Project Renaissance Analysis - AU", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")


def extract_cc(cost_centre):
    return cost_centre[:9]


def extract_site(value):
    return value.split(" - ")[1].lower().title()


def extract_name(value):
    return value.split(" [")[0].lower().title()


def extract_cost_name(value):
    if " AU - " in value:
        return value.split(" AU - ")[1].lower().title()
    elif " NZ - " in value:
        return value.split(" NZ - ")[1].lower().title()


def format_for_percentage(value):
    return "{:.2%}".format(value)


def format_for_currency(value):
    return "${0:,.0f}".format(value)


def format_for_float_currency(value):
    return "${0:,.2f}".format(value)


def format_for_float(value):
    return "{0:,.2f}".format(value)


def format_for_int(value):
    return "{0:,.0f}".format(value)


def proper_case(value):
    return value.title()


def extract_short_name(value):
    firstname = value.split(" [")[0].lower().title()
    if " " in firstname:
        second_name = firstname.split(" ")[1][:4].title()
    else:
        second_name = " "

    firstname = firstname.split(" ")[0].title()

    return firstname + " " + second_name


today = datetime.datetime.now()


def date_three_weeks_ago(today_date, number_of_weeks=3):
    # Subtract 3 weeks (21 days) from the given date
    three_weeks_ago = today_date - timedelta(weeks=number_of_weeks)
    return three_weeks_ago


# Side for Uploading excel File ###############################

with st.spinner('Loading and Updating Report...🥱'):
    try:

        with st.sidebar:

            st.title("Analytics Dashboard")
            uploaded_file = st.file_uploader('Upload Customer Profitability Summary File',
                                             type=['xlsx', 'xlsm', 'xlsb'])
            st.divider()

            st.title("Customer Rates Summary File")
            customer_rates_file = st.file_uploader('Upload Standardisation File', type=['xlsx', 'xlsm', 'xlsb'])
            st.markdown(f'###### This is the Customer Rates Model from Box:APAC Rev...')
            st.divider()

            st.title("Site Detailed View")
            uploaded_invoicing_data = st.file_uploader('Upload Customer Rates File', type=['xlsx', 'xlsm', 'xlsb'])
            invoice_rates = load_invoices_model(uploaded_invoicing_data)
            # Main Dashboard

    except Exception as e:

        st.subheader('To use with this WebApp 📊  - Upload below : \n '
                     '1. Profitability Summary File \n'
                     '2. Customer Rates Summary File \n'
                     '3. Customer Invoicing Data')

# Main Dashboard ##############################


if uploaded_file and customer_rates_file and uploaded_invoicing_data:

    invoice_rates = invoice_rates

    DetailsTab, AboutTab = st.tabs(["📊 Quick View",
                                    # "🥇 Site Detailed View",
                                    # "📈 Estimate Profitability",
                                    "ℹ️ About"])


    def highlight_negative_values(value):
        if type(value) != str:
            if value < 0:
                return f"color: red"


    profitability_summary_file = load_profitbility_Summary_model(uploaded_file)

    # Create additional Column called CC : to be used for cross-referencing with Customer Rates Fil
    profitability_summary_file["CC"] = profitability_summary_file["Cost Center"].apply(extract_cc)  # Add CC to be

    # Create additional Column called Name : to be used for displaying friendly Customer Name
    profitability_summary_file["Name"] = profitability_summary_file["Customer"].apply(extract_name)

    # Create additional Column called Name : to be used for displaying friendly Customer Name
    profitability_summary_file["Site"] = profitability_summary_file["Cost Center"].apply(extract_site)

    # Read individual Excel sheet separately = Prior Year Customer Profitability. Use for comparison.
    profitability_2023 = load_specific_xls_sheet(customer_rates_file, "2023_Profitability", 0, "A:AQ")
    profitability_2023["Site"] = profitability_2023["Cost Center"].apply(extract_site)

    # Read individual Excel sheet separately = Customer Volumes current year and prior year. Use for comparison.
    customer_pallets = load_specific_xls_sheet(customer_rates_file, "customer_pallets", 0, "A:CV")
    customer_pallets["Site"] = customer_pallets["Cost Center"].apply(extract_site)

    # Read individual Excel sheet separately = Customer Volumes current year and prior year. Use for comparison.
    customer_rate_cards = load_specific_xls_sheet(customer_rates_file, "2025_Rate_Cards", 0, "A:J")

    # Read individual Excel sheet separately = Customer Volumes current year and prior year. Use for comparison.
    dominic_report = load_specific_xls_sheet(customer_rates_file, "dominic_report", 0, "A:AM")

    # Read individual Excel sheet separately = Customer Volumes current year and prior year. Use for comparison.
    budget_data_2025 = load_specific_xls_sheet(customer_rates_file, "2025_Budget", 0, "A:AP")

    budget_data_2025 = budget_data_2025[["Year", "Site", "Revenue", "Ebitda", "Economic OHP", "Labour to Tot. Rev",
                                         "Direct Labor / Hour", "DL to Svcs Rev", "Economic Utilization", "Services",
                                         "Throughput Plt", "Total Pallets", "Physical OHP", "Rent & Storage & Blast"
                                         ]]

    budget_data_2025["Rev Per Pallet"] = budget_data_2025["Revenue"] / ((budget_data_2025["Economic OHP"] / 12) * 52)
    budget_data_2025["Ebitda Per Pallet"] = budget_data_2025["Ebitda"] / (
        ((budget_data_2025["Economic OHP"] / 12) * 52))
    budget_data_2025["Turn"] = ((budget_data_2025["Throughput Plt"] / 2) / (budget_data_2025["Physical OHP"] / 12))

    budget_data_2025["Rev - Storage & Blast"] = budget_data_2025["Rent & Storage & Blast"] / (
            budget_data_2025["Rent & Storage & Blast"] + budget_data_2025["Services"])

    budget_data_2025["Rev - Services"] = budget_data_2025["Services"] / (
            budget_data_2025["Rent & Storage & Blast"] + budget_data_2025["Services"])
    budget_data_2025.dropna(inplace=True)
    budget_data_2025.drop(
        columns=["Throughput Plt", "Total Pallets", "Physical OHP", "Services", "Rent & Storage & Blast"], axis=1,
        inplace=True)
    budget_comparison_years = budget_data_2025.Year.unique()

    # Customer Names for Selection in Select Box
    all_workday_customer_names = profitability_summary_file.Customer.unique()

    # Site Names for Selection in Select Box
    country_region = list(profitability_summary_file.Region.unique())

    #  Details Tab ###############################

    with (DetailsTab):
        title_holder, country_region_selection, cost_centres_selection = st.columns((1, 1, 3))

        title_holder.subheader("Project Renaissance", divider="blue")
        # selected_cost_centre = st.selectbox("Select Cost Centre :", cost_centres, index=0, )
        selected_region = country_region_selection.multiselect('Region :', country_region, country_region[0], key=203)
        selected_region_profitability = profitability_summary_file[
            profitability_summary_file["Region"].isin(selected_region)]
        site_list = list(selected_region_profitability.Site.unique())
        selected_site = cost_centres_selection.multiselect("Site :", site_list, site_list, key=206)
        # countries = ["Australia", "New Zealand"]
        # country_selected = country_label.selectbox("Country", countries, index=0, key=189)

        # Cost Centre Names for Selection in Select Box
        cost_centres = profitability_summary_file[profitability_summary_file["Site"].isin(selected_site)]
        cost_centres = cost_centres["Cost Center"].unique()

        st.text("")
        st.text("")
        st.text("")

        select_Site_data = profitability_summary_file[profitability_summary_file.Site.isin(selected_site)]

        select_Site_data_2023_profitability = profitability_2023[
            profitability_2023.Site.isin(selected_site)]
        select_Site_data_2023_pallets = customer_pallets[customer_pallets.Site.isin(selected_site)]

        fin1, fin2 = st.columns((2.5, 3))

        key_metrics_data = profitability_summary_file[profitability_summary_file[" Revenue"] > 1]
        key_metrics_data = key_metrics_data[key_metrics_data.Site.isin(selected_site)]
        key_metrics_data["Name"] = key_metrics_data["Cost Center"].apply(extract_cost_name)

        # key_metrics_data["Sqm Rent"] = key_metrics_data['Rent Expense,\n$'] / key_metrics_data['sqm']

        key_metrics_data_pivot = pd.pivot_table(key_metrics_data, index=["Cost Center", "Name", "Site"],
                                                values=[" Revenue", 'EBITDAR $', 'Rent Expense,\n$',
                                                        'EBITDA $',
                                                        # 'Sqm Rent'
                                                        ],
                                                aggfunc='sum', sort=True)

        key_metrics_data_pivot_Hume = pd.pivot_table(key_metrics_data, index=["Site"],
                                                     values=[" Revenue", 'EBITDAR $', 'Rent Expense,\n$',
                                                             'EBITDA $',
                                                             # 'Sqm Rent'
                                                             ],
                                                     aggfunc='sum', sort=True)

        key_metrics_data_pivot = key_metrics_data_pivot.reset_index("Name")

        key_metrics_data_pivot["Ebitda - %"] = key_metrics_data_pivot['EBITDA $'] / key_metrics_data_pivot[' Revenue']
        key_metrics_data_pivot.insert(3, "Ebitdar - %",
                                      key_metrics_data_pivot['EBITDAR $'] / key_metrics_data_pivot[' Revenue'])

        key_metrics_data_pivot_Hume["Ebitda - %"] = key_metrics_data_pivot_Hume['EBITDA $'] / \
                                                    key_metrics_data_pivot_Hume[' Revenue']

        key_metrics_data_pivot_Hume.insert(2, "Ebitdar - %", key_metrics_data_pivot_Hume['EBITDAR $'] / \
                                           key_metrics_data_pivot_Hume[' Revenue'])

        key_metrics_data_pivot = key_metrics_data_pivot.style.hide(axis="index")

        key_metrics_data_pivot = key_metrics_data_pivot.format({
            " Revenue": '${0:,.0f}',
            'EBITDAR $': "${0:,.0f}",
            "Ebitdar - %": "{0:,.2%}",
            'Rent Expense,\n$': "${0:,.0f}",
            'EBITDA $': "${0:,.0f}",
            'Ebitda - %': "{0:,.2%}",
        })

        key_metrics_data_pivot = key_metrics_data_pivot.map(highlight_negative_values)

        key_metrics_data_pivot = style_dataframe(key_metrics_data_pivot)

        # Total for Hume Road

        key_metrics_data_pivot_Hume = key_metrics_data_pivot_Hume.reset_index("Site")
        key_metrics_data_pivot_Hume.rename(columns={"Site": "Site   Total  for All Building Hume"}, inplace=True)

        key_metrics_data_pivot_Hume = key_metrics_data_pivot_Hume.style.hide(axis="index")

        key_metrics_data_pivot_Hume = key_metrics_data_pivot_Hume.format({
            " Revenue": '${0:,.0f}',
            'EBITDAR $': "${0:,.0f}",
            "Ebitdar - %": "{0:,.2%}",
            'Rent Expense,\n$': "${0:,.0f}",
            'EBITDA $': "${0:,.0f}",
            'Ebitda - %': "{0:,.2%}",
        })

        key_metrics_data_pivot_Hume = key_metrics_data_pivot_Hume.map(highlight_negative_values)

        key_metrics_data_pivot_Hume = style_dataframe(key_metrics_data_pivot_Hume)

        with fin1:
            fin1.markdown("##### Financial Summary")
            fin1.write(key_metrics_data_pivot.to_html(), unsafe_allow_html=True, )
            if "Hume Rd" == selected_site[0]:
                st.text("")
                st.text("")
                st.text("")
                fin1.write(key_metrics_data_pivot_Hume.to_html(), unsafe_allow_html=True, )

        key_metrics_data_Hume_KPI = key_metrics_data.groupby("Site").sum()

        test = key_metrics_data[["CC", "Pal Cap"]].groupby("CC").mean()
        key_metrics_data = key_metrics_data.groupby("CC").sum()

        DL_Handling = key_metrics_data["Direct Labor Expense,\n$"]
        Ttl_Labour = key_metrics_data["Total Labor Expense, $"]
        Service_Revenue = key_metrics_data["Service Rev (Handling+ Case Pick+ Other Rev)"]

        key_metrics_data["Capacity"] = key_metrics_data["Pallet"]
        key_metrics_data["Pallets"] = key_metrics_data["Pallet"]
        key_metrics_data["Occ-%"] = key_metrics_data["Pallet"] / key_metrics_data["Capacity"]
        key_metrics_data["Rev | Plt"] = key_metrics_data[" Revenue"] / (key_metrics_data["Pallet"] * 52)
        key_metrics_data["EBITDA | Plt"] = key_metrics_data["EBITDA $"] / (key_metrics_data["Pallet"] * 52)
        key_metrics_data["Rent | Plt"] = (key_metrics_data["Rent Expense,\n$"] / key_metrics_data["Pallet"]) / 52
        key_metrics_data["Turn"] = ((key_metrics_data["TTP p.w."] / 2) * 52) / key_metrics_data["Pallet"]
        key_metrics_data["DL %"] = DL_Handling / Service_Revenue
        key_metrics_data["LTR %"] = Ttl_Labour / key_metrics_data[" Revenue"]
        key_metrics_data["Rev | SQM"] = key_metrics_data[" Revenue"] / key_metrics_data["sqm"]
        key_metrics_data["EBITDA | SQM"] = key_metrics_data["EBITDA $"] / key_metrics_data["sqm"]
        key_metrics_data["Rent | SQM"] = key_metrics_data['Rent Expense,\n$'] / key_metrics_data['sqm']

        # key_metrics_data["Site Pal Cap psqm"] = key_metrics_data['Site Pal Cap psqm']

        kpi_key_metrics_data_pivot = key_metrics_data[
            ["Pallet",
             "Capacity",
             "Pallets", "Occ-%", "Rev | Plt", "Rent | Plt", "EBITDA | Plt", "Turn", "DL %", "LTR %", "Rev | SQM",
             "Rent | SQM",
             "EBITDA | SQM",

             ]]

        kpi_key_metrics_data_pivot = pd.concat([kpi_key_metrics_data_pivot, test], axis=1, )
        kpi_key_metrics_data_pivot["Capacity"] = kpi_key_metrics_data_pivot["Pal Cap"]
        kpi_key_metrics_data_pivot["Occ-%"] = kpi_key_metrics_data_pivot["Pallet"] / kpi_key_metrics_data_pivot[
            "Capacity"]

        budget_data_2025_bench_mark = kpi_key_metrics_data_pivot

        kpi_key_metrics_data_pivot = kpi_key_metrics_data_pivot.style.hide(axis="index")
        kpi_key_metrics_data_pivot = kpi_key_metrics_data_pivot.hide(["Pallet", "Pal Cap"], axis="columns")

        kpi_key_metrics_data_pivot = kpi_key_metrics_data_pivot.format({
            'Rev | Plt': "${0:,.2f}",
            'EBITDA | Plt': "${0:,.2f}",
            'Turn': "{0:,.2f}",
            'DL %': '{0:,.1%}',
            'LTR %': '{0:,.1%}',
            'Rev | SQM': "${0:,.2f}",
            'EBITDA | SQM': "${0:,.2f}",
            'Rent | SQM': "${0:,.2f}",
            "Capacity": "{0:,.0f}",
            "Pallets": "{0:,.0f}",
            "Occ-%": '{0:,.1%}',
            "Rent | Plt": '${0:,.2f}',
        })

        kpi_key_metrics_data_pivot = kpi_key_metrics_data_pivot.map(highlight_negative_values)

        kpi_key_metrics_data_pivot = style_dataframe(kpi_key_metrics_data_pivot)

        with fin2:
            fin2.markdown("##### KPI Summary")
            fin2.write(kpi_key_metrics_data_pivot.to_html(), unsafe_allow_html=True)

            DL_Handling = key_metrics_data_Hume_KPI["Direct Labor Expense,\n$"]
            Ttl_Labour = key_metrics_data_Hume_KPI["Total Labor Expense, $"]
            Service_Revenue = key_metrics_data_Hume_KPI["Service Rev (Handling+ Case Pick+ Other Rev)"]

            # key_metrics_data["Capacity"] =   key_metrics_data["CC Pal Cap"]
            key_metrics_data_Hume_KPI["Capacity"] = 155737
            key_metrics_data_Hume_KPI["Pallets"] = key_metrics_data_Hume_KPI["Pallet"]
            key_metrics_data_Hume_KPI["Occ-%"] = key_metrics_data_Hume_KPI["Pallet"] / key_metrics_data_Hume_KPI[
                "Capacity"]
            key_metrics_data_Hume_KPI["Rev | Plt"] = key_metrics_data_Hume_KPI[" Revenue"] / (
                    key_metrics_data_Hume_KPI["Pallet"] * 52)
            key_metrics_data_Hume_KPI["EBITDA | Plt"] = key_metrics_data_Hume_KPI["EBITDA $"] / (
                    key_metrics_data_Hume_KPI["Pallet"] * 52)
            key_metrics_data_Hume_KPI["Rent | Plt"] = (key_metrics_data_Hume_KPI["Rent Expense,\n$"] /
                                                       key_metrics_data_Hume_KPI["Pallet"]) / 52
            key_metrics_data_Hume_KPI["Turn"] = ((key_metrics_data_Hume_KPI["TTP p.w."] / 2) * 52) / \
                                                key_metrics_data_Hume_KPI["Pallet"]
            key_metrics_data_Hume_KPI["DL %"] = DL_Handling / Service_Revenue
            key_metrics_data_Hume_KPI["LTR %"] = Ttl_Labour / key_metrics_data_Hume_KPI[" Revenue"]
            key_metrics_data_Hume_KPI["Rev | SQM"] = key_metrics_data_Hume_KPI[" Revenue"] / key_metrics_data_Hume_KPI[
                "sqm"]
            key_metrics_data_Hume_KPI["EBITDA | SQM"] = key_metrics_data_Hume_KPI["EBITDA $"] / \
                                                        key_metrics_data_Hume_KPI["sqm"]
            key_metrics_data_Hume_KPI["Rent | SQM"] = key_metrics_data_Hume_KPI['Rent Expense,\n$'] / \
                                                      key_metrics_data_Hume_KPI['sqm']
            # key_metrics_data["Site Pal Cap psqm"] = key_metrics_data['Site Pal Cap psqm']

            kpi_key_metrics_data_pivot_Hume_KPI = key_metrics_data_Hume_KPI[
                [
                    "Capacity",
                    "Pallets", "Occ-%", "Rev | Plt", "Rent | Plt", "EBITDA | Plt", "Turn", "DL %", "LTR %", "Rev | SQM",
                    "EBITDA | SQM",
                ]]

            kpi_key_metrics_data_pivot_Hume_KPI = kpi_key_metrics_data_pivot_Hume_KPI.style.hide(axis="index")

            kpi_key_metrics_data_pivot_Hume_KPI = kpi_key_metrics_data_pivot_Hume_KPI.format({
                'Rev | Plt': "${0:,.2f}",
                'EBITDA | Plt': "${0:,.2f}",
                'Turn': "{0:,.2f}",
                'DL %': '{0:,.1%}',
                'LTR %': '{0:,.1%}',
                'Rev | SQM': "${0:,.2f}",
                'EBITDA | SQM': "${0:,.2f}",
                'Rent | SQM': "${0:,.2f}",
                "Capacity": "{0:,.0f}",
                "Pallets": "{0:,.0f}",
                "Occ-%": '{0:,.1%}',
                "Rent | Plt": '${0:,.2f}',
            })

            kpi_key_metrics_data_pivot_Hume_KPI = kpi_key_metrics_data_pivot_Hume_KPI.map(highlight_negative_values)

            kpi_key_metrics_data_pivot_Hume_KPI = style_dataframe(kpi_key_metrics_data_pivot_Hume_KPI)
            st.text("")
            st.text("")
            st.text("")
            if "Hume Rd" == selected_site[0]:
                fin2.write(kpi_key_metrics_data_pivot_Hume_KPI.to_html(), unsafe_allow_html=True)

        # TOP 10 Customers ##################################################

        with st.expander("Top Customers Review", expanded=True):

            profitability_by_Customer, customer_network_profitability, commodity_type_profitability = st.tabs(
                ["📊 Top Site Customers",
                 "🥇 Multi-Site Customer Profitability View",
                 "🧺🧺 Commodity Mix Profitability"
                 ])

            with(profitability_by_Customer):
                st.subheader("Top Customers", divider='rainbow')

                st.text("")

                s1, s2, s3, s4, s5 = st.columns(5)

                with s1:
                    selected_cost_centre = st.selectbox("Select Cost Centre :", cost_centres, index=0, )

                selected_CC = selected_cost_centre[:9]
                select_CC_data = profitability_summary_file[profitability_summary_file["CC"] == selected_CC]
                # select_CC_data_2023_profitability = profitability_2023[
                #     profitability_2023["CC"] == selected_CC]
                # select_CC_data_2023_pallets = customer_pallets[customer_pallets["CC"] == selected_CC]

                select_CC_data = select_CC_data.query("Customer != '.All Other [.All Other]'")
                select_CC_data = select_CC_data.query("Pallet > 1")
                select_CC_data["Rank"] = select_CC_data[" Revenue"].rank(method='max')
                select_CC_data["Rev psqm"] = select_CC_data[" Revenue"] / select_CC_data["sqm"]
                treemap_data = select_CC_data.filter(
                    ['Name', " Revenue", 'EBITDAR $', 'EBITDAR Margin\n%', 'EBITDA $', 'EBITDA Margin\n%',
                     'Revenue per Pallet', 'EBITDA per Pallet',
                     'Pallet', 'TTP p.w.', 'Rank', 'Turn', "DLH% \n(DL / Service Rev)", "LTR % (Labour to Rev %)",
                     'Rev psqm',
                     'EBITDA psqm', 'sqm', 'Rent psqm', 'Storage Revenue, $', 'Blast Freezing Revenue, $',
                     'RSB Rev per OHP',
                     'Service Rev per TPP', 'Rent Expense,\n$'
                     ])

                # treemap_data.rename(columns={'EBITDAR Margin\n%': 'EBITDAR %','EBITDA Margin\n%':'EBITDA %','EBITDA per Pallet':'EBITDA | Plt',
                #                              'Revenue per Pallet': 'Rev | Plt','DLH% \n(DL / Service Rev)': 'DL Ratio' })

                treemap_data.insert(7, "Rent | Plt", (treemap_data['Rent Expense,\n$'] / (treemap_data.Pallet * 52)))

                display_data = treemap_data
                size = len(display_data)
                customer_view_size = s2.number_input("Filter Bottom Outlier Customers", value=size, key=772)
                display_data = display_data.query(f"Rank > {size - customer_view_size} ")
                display_data_pie = display_data
                rank_display_data = display_data
                score_carding_data = display_data
                display_data["RSB"] = (display_data['Storage Revenue, $'] + display_data[
                    'Blast Freezing Revenue, $']) / display_data[" Revenue"]
                display_data["Services"] = 1 - display_data["RSB"]
                display_data.rename(
                    columns={"DLH% \n(DL / Service Rev)": "DL ratio", "LTR % (Labour to Rev %)": "LTR - %",
                             'EBITDAR Margin\n%': 'EBITDAR %', 'EBITDA Margin\n%': 'EBITDA %',
                             'EBITDA per Pallet': 'EBITDA | Plt',
                             'Revenue per Pallet': 'Rev | Plt'}, inplace=True)

                treemap_graph_data = display_data
                display_data = display_data.style.hide(axis="index")

                display_data = display_data.hide(
                    ['Rank', 'sqm', 'Storage Revenue, $', 'Blast Freezing Revenue, $', 'Rent psqm',
                     'Rent Expense,\n$',
                     ],
                    axis="columns")

                display_data = style_dataframe(display_data)

                display_data = display_data.format({
                    " Revenue": '${0:,.0f}',
                    'EBITDAR $': "${0:,.0f}",
                    'EBITDAR %': "{0:,.2%}",
                    'EBITDA $': "${0:,.0f}",
                    'EBITDA %': "{0:,.2%}",
                    'Rev | Plt': "${0:,.2f}",
                    'EBITDA | Plt': "${0:,.2f}",
                    'Pallet': "{0:,.0f}",
                    'TTP p.w.': '{0:,.0f}',
                    "DL ratio": "{0:,.2%}",
                    "LTR - %": "{0:,.2%}",
                    'Rev psqm': "${0:,.2f}",
                    'EBITDA psqm': "${0:,.2f}",
                    "Rent | Plt": "${0:,.2f}",
                    'Rent psqm': "${0:,.2f}",
                    'Turn': "{0:,.2f}",
                    'RSB': "{0:,.2%}",
                    'Services': "{0:,.2%}",
                    'RSB Rev per OHP': "${0:,.2f}",
                    'Service Rev per TPP': "${0:,.2f}",

                })

                display_data = display_data.map(highlight_negative_values)
                display_data.background_gradient(subset=['EBITDA %'], cmap="RdYlGn")

                # Create a download button
                with s3:
                    s3.text("")
                    output1 = io.BytesIO()
                    with pd.ExcelWriter(output1) as writer:
                        display_data.to_excel(writer, sheet_name='export_data', index=False)

                    st.text("")
                    s3.download_button(
                        label="👆 Download ⤵️",
                        data=output1,
                        file_name='customer_revenue.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=530,
                    )

                st.write(display_data.to_html(), unsafe_allow_html=True)
                # st.write(display_data.to_html(),   unsafe_allow_html=True, use_container_width=True )
                st.text("")
                st.text("")
                st.text("")

            with(customer_network_profitability):
                st.subheader("Multi-Site Customer Profitability", divider='rainbow')

                st.text("")
                st.text("")

                s1, s2, s3 = st.columns(3)

                with s1:
                    profitability_summary_file = profitability_summary_file.query(
                        "Customer != '.All Other [.All Other]'")

                    multi_site_customers = profitability_summary_file.sort_values(by="Name").Name.unique()
                    default_customers = [multi_site_customers[8], multi_site_customers[27], multi_site_customers[28]]

                    select_network_customers = st.multiselect("Select Customers : ", multi_site_customers,
                                                              default_customers, key=792)

                select_network_data = profitability_summary_file[
                    profitability_summary_file.Name.isin(select_network_customers)]

                select_network_data = select_network_data.query("Pallet > 1")
                select_network_data["Rank"] = select_network_data[" Revenue"].rank(method='max')
                select_network_data["Rev psqm"] = select_network_data[" Revenue"] / select_network_data["sqm"]

                treemap_network_data = select_network_data.filter(
                    ['Name', "Site", "Cost Center", " Revenue", 'EBITDAR $', 'EBITDAR Margin\n%', 'EBITDA $',
                     'EBITDA Margin\n%', 'Revenue per Pallet', 'EBITDA per Pallet',
                     'Pallet', 'TTP p.w.', 'Rank', 'Turn', "DLH% \n(DL / Service Rev)", "LTR % (Labour to Rev %)",
                     'Rev psqm',
                     'EBITDA psqm', 'sqm', 'Rent psqm', 'Storage Revenue, $', 'Blast Freezing Revenue, $',
                     'RSB Rev per OHP',
                     'Service Rev per TPP', 'Rent Expense,\n$', 'Avg. Storage Rate (Storage $/Pallet)',
                     'Avg. Handling Rate (Handling $/TTP)'
                     ])
                treemap_network_data["Cost Center"] = [x.split(" - ")[1] + " - " + x.split(" - ")[2] for x in
                                                       treemap_network_data["Cost Center"]]

                treemap_network_data.insert(7, "Rent | Plt", (
                        treemap_network_data['Rent Expense,\n$'] / (treemap_network_data.Pallet * 52)))
                treemap_network_data = treemap_network_data.sort_values(by=" Revenue", ascending=False)

                display_network_data = treemap_network_data
                size = len(display_network_data)

                network_customer_view_size = s2.number_input("Filter Bottom Outlier Customers", value=size, key=887)
                display_network_data = display_network_data.query(f"Rank > {size - network_customer_view_size} ")

                display_network_data["RSB"] = (display_network_data['Storage Revenue, $'] + display_network_data[
                    'Blast Freezing Revenue, $']) / display_network_data[" Revenue"]
                display_network_data["Services"] = 1 - display_network_data["RSB"]
                display_network_data.rename(
                    columns={"DLH% \n(DL / Service Rev)": "DL ratio", "LTR % (Labour to Rev %)": "LTR - %",
                             'EBITDAR Margin\n%': 'EBITDAR %', 'EBITDA Margin\n%': 'EBITDA %',
                             'EBITDA per Pallet': 'EBITDA | Plt',
                             'Revenue per Pallet': 'Rev | Plt',
                             'Avg. Storage Rate (Storage $/Pallet)': 'Avg. Storage Rate',
                             'Avg. Handling Rate (Handling $/TTP)': 'Avg. Handling Rate'

                             }, inplace=True)

                # treemap_graph_data = display_data
                display_network_data = display_network_data.style.hide(axis="index")

                display_network_data = display_network_data.hide(
                    ['Rank', 'sqm', 'Storage Revenue, $', 'Blast Freezing Revenue, $', 'Rent psqm',
                     'Rent Expense,\n$', 'Rev psqm', 'EBITDA psqm', 'Avg. Storage Rate', 'Site',
                     'Avg. Handling Rate'
                     ],
                    axis="columns")

                display_network_data = style_dataframe(display_network_data)

                display_network_data = display_network_data.format({
                    " Revenue": '${0:,.0f}',
                    'EBITDAR $': "${0:,.0f}",
                    'EBITDAR %': "{0:,.2%}",
                    'EBITDA $': "${0:,.0f}",
                    'EBITDA %': "{0:,.2%}",
                    'Rev | Plt': "${0:,.2f}",
                    'EBITDA | Plt': "${0:,.2f}",
                    'Pallet': "{0:,.0f}",
                    'TTP p.w.': '{0:,.0f}',
                    "DL ratio": "{0:,.2%}",
                    "LTR - %": "{0:,.2%}",
                    'Rev psqm': "${0:,.2f}",
                    'EBITDA psqm': "${0:,.2f}",
                    "Rent | Plt": "${0:,.2f}",
                    'Rent psqm': "${0:,.2f}",
                    'Turn': "{0:,.2f}",
                    'RSB': "{0:,.2%}",
                    'Services': "{0:,.2%}",
                    'RSB Rev per OHP': "${0:,.2f}",
                    'Service Rev per TPP': "${0:,.2f}",

                })

                display_network_data = display_network_data.map(highlight_negative_values)
                display_network_data.background_gradient(subset=['EBITDA %'], cmap="RdYlGn")

                # Create a download button
                with s3:
                    s3.text("")
                    output863 = io.BytesIO()
                    with pd.ExcelWriter(output863) as writer:
                        display_network_data.to_excel(writer, sheet_name='export_data', index=False)

                    st.text("")
                    s3.download_button(
                        label="👆 Download ⤵️",
                        data=output863,
                        file_name='customer_revenue.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=877,
                    )

                st.write(display_network_data.to_html(), unsafe_allow_html=True, )
                # st.write(display_network_data.to_html(), unsafe_allow_html=True, use_container_width=True)

                st.text("")
                st.text("")
                st.text("")

            with(commodity_type_profitability):
                st.subheader("Commodity Mix", divider='rainbow')

                st.text("")

                ct1, ct2, ct3, ct4, ct5 = st.columns((1, 3, 1, 1, 0.5))
                # Set the Region to view
                commodity_type_region = sorted(profitability_summary_file.Region.unique())
                commodity_type_region = np.insert(commodity_type_region, len(commodity_type_region), "Country")
                selected_region_commodity_type = ct1.multiselect("Select Region :", commodity_type_region,
                                                                 commodity_type_region[0], key=981)

                if selected_region_commodity_type[0] == "Country":
                    commodity_type_profitability_data = profitability_summary_file
                else:
                    # Filter the data for set Region
                    commodity_type_profitability_data = profitability_summary_file[
                        profitability_summary_file["Region"].isin(selected_region_commodity_type)]

                # Set Cost centre to View
                site_list_commodity_type = list(commodity_type_profitability_data.Site.unique())

                with ct2:
                    selected_site_commodity_type = st.multiselect("Select Cost Centre :", site_list_commodity_type,
                                                                  site_list_commodity_type, key=977)

                cost_centres_commodity_type = profitability_summary_file[
                    profitability_summary_file['Site'].isin(selected_site_commodity_type)]
                cost_centres_commodity_type = cost_centres_commodity_type["Cost Center"].unique()

                # Set Cost Centre Dataa
                commodity_type_cost_centres_data = commodity_type_profitability_data[
                    commodity_type_profitability_data["Site"].isin(selected_site_commodity_type)]

                # selected_CC = selected_cost_centre[:9]

                # select_CC_data = commodity_type_cost_centres_data[commodity_type_cost_centres_data["CC"] == selected_cost_centre]
                select_CC_data = commodity_type_cost_centres_data.groupby("Commodity Type").sum()
                select_CC_data = select_CC_data.reset_index("Commodity Type")

                # select_CC_data_2023_profitability = profitability_2023[
                #     profitability_2023["CC"] == selected_CC]
                # select_CC_data_2023_pallets = customer_pallets[customer_pallets["CC"] == selected_CC]

                select_CC_data = select_CC_data.query("Customer != '.All Other [.All Other]'")
                select_CC_data = select_CC_data.query("Pallet > 1")
                select_CC_data["Rank"] = select_CC_data[" Revenue"].rank(method='min')
                select_CC_data = select_CC_data.sort_values(by=" Revenue", ascending=False)

                # treemap_network_data = treemap_network_data.sort_values(by=" Revenue", ascending=False)

                select_CC_data["Rev psqm"] = select_CC_data[" Revenue"] / select_CC_data["sqm"]
                select_CC_data["EBITDA psqm"] = select_CC_data['EBITDA $'] / select_CC_data["sqm"]

                treemap_data = select_CC_data.filter(
                    ['Commodity Type', "Name", " Revenue", 'EBITDAR $', 'EBITDAR Margin\n%', 'EBITDA $',
                     'EBITDA Margin\n%',
                     'Revenue per Pallet', 'EBITDA per Pallet',
                     'Pallet', 'TTP p.w.', 'Rank', 'Turn', "DLH% \n(DL / Service Rev)", "LTR % (Labour to Rev %)",
                     'Rev psqm',
                     'EBITDA psqm', 'sqm', 'Rent psqm', 'Storage Revenue, $', 'Blast Freezing Revenue, $',
                     'RSB Rev per OHP',
                     'Service Rev per TPP', 'Rent Expense,\n$',
                     "Direct Labor Expense,\n$",
                     "Total Labor Expense, $",
                     "Service Rev (Handling+ Case Pick+ Other Rev)",
                     ])

                treemap_data.insert(7, "Rent | Plt", (treemap_data['Rent Expense,\n$'] / (treemap_data.Pallet * 52)))

                display_data = treemap_data
                size = len(display_data)
                customer_view_size = ct3.number_input("Filter Bottom Outlier Customers", value=size, key=1012)
                display_data = display_data.query(f"Rank > {size - customer_view_size} ")
                display_data_pie = display_data
                rank_display_data = display_data
                score_carding_data = display_data
                display_data["RSB"] = (display_data['Storage Revenue, $'] + display_data[
                    'Blast Freezing Revenue, $']) / display_data[" Revenue"]
                display_data["Services"] = 1 - display_data["RSB"]
                display_data.rename(
                    columns={"DLH% \n(DL / Service Rev)": "DL ratio", "LTR % (Labour to Rev %)": "LTR - %",
                             'EBITDAR Margin\n%': 'EBITDAR %', 'EBITDA Margin\n%': 'EBITDA %',
                             'EBITDA per Pallet': 'EBITDA | Plt',
                             'Revenue per Pallet': 'Rev | Plt'}, inplace=True)

                commodity_type_DL_Handling = display_data["Direct Labor Expense,\n$"]
                commodity_type_ttl_Labour = display_data["Total Labor Expense, $"]
                commodity_type_service_Revenue = display_data["Service Rev (Handling+ Case Pick+ Other Rev)"]

                display_data["EBITDA %"] = display_data['EBITDA $'] / display_data[' Revenue']
                display_data["EBITDAR %"] = display_data['EBITDAR $'] / display_data[' Revenue']
                display_data["Rev | Plt"] = display_data[' Revenue'] / (display_data['Pallet'] * 52)
                display_data["EBITDA | Plt"] = display_data['EBITDA $'] / (display_data['Pallet'] * 52)
                display_data["Turn"] = ((display_data["TTP p.w."] / 2) * 52) / display_data["Pallet"]
                display_data["DL ratio"] = commodity_type_DL_Handling / commodity_type_service_Revenue
                display_data["LTR - %"] = commodity_type_ttl_Labour / display_data[" Revenue"]
                display_data['RSB Rev per OHP'] = (display_data['Storage Revenue, $'] + display_data[
                    'Blast Freezing Revenue, $']) / (display_data['Pallet'] * 52)
                display_data['Service Rev per TPP'] =( display_data[" Revenue"]-  (display_data['Storage Revenue, $'] + display_data[
                    'Blast Freezing Revenue, $']) ) / (display_data['TTP p.w.'] * 52)

                treemap_graph_data = display_data

                display_data = display_data.style.hide(axis="index")

                display_data = display_data.hide(
                    ["Name", 'Rank', 'sqm', 'Storage Revenue, $',
                     'Blast Freezing Revenue, $',
                     'Rent psqm',
                     'Rent Expense,\n$',
                     "Direct Labor Expense,\n$",
                     "Total Labor Expense, $",
                     "Service Rev (Handling+ Case Pick+ Other Rev)",
                     ],
                    axis="columns")

                display_data = style_dataframe(display_data)

                display_data = display_data.format({
                    " Revenue": '${0:,.0f}',
                    'EBITDAR $': "${0:,.0f}",
                    'EBITDAR %': "{0:,.2%}",
                    'EBITDA $': "${0:,.0f}",
                    'EBITDA %': "{0:,.2%}",
                    'Rev | Plt': "${0:,.2f}",
                    'EBITDA | Plt': "${0:,.2f}",
                    'Pallet': "{0:,.0f}",
                    'TTP p.w.': '{0:,.0f}',
                    "DL ratio": "{0:,.2%}",
                    "LTR - %": "{0:,.2%}",
                    'Rev psqm': "${0:,.2f}",
                    'EBITDA psqm': "${0:,.2f}",
                    "Rent | Plt": "${0:,.2f}",
                    'Rent psqm': "${0:,.2f}",
                    'Turn': "{0:,.2f}",
                    'RSB': "{0:,.2%}",
                    'Services': "{0:,.2%}",
                    'RSB Rev per OHP': "${0:,.2f}",
                    'Service Rev per TPP': "${0:,.2f}",

                })

                display_data = display_data.map(highlight_negative_values)
                display_data.background_gradient(subset=['EBITDA %'], cmap="RdYlGn")

                # Create a download button
                with ct4:
                    ct4.text("")
                    output1 = io.BytesIO()
                    with pd.ExcelWriter(output1) as writer:
                        display_data.to_excel(writer, sheet_name='export_data', index=False)

                    st.text("")
                    ct4.download_button(
                        label="👆 Download ⤵️",
                        data=output1,
                        file_name='customer_revenue.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=1071,
                    )

                st.write(display_data.to_html(), unsafe_allow_html=True)
                st.text("")
                st.text("")
                st.text("")

                customer_names_display = commodity_type_cost_centres_data

                displayCustomers_map = {}

                for commodityType in customer_names_display["Commodity Type"].unique():
                    displayCustomers = sorted(
                        customer_names_display[customer_names_display['Commodity Type'] == commodityType][
                            "Name"].unique())

                    displayCustomers_map[commodityType] = ' | '.join(displayCustomers)

                displayCustomers_result = pd.DataFrame(list(displayCustomers_map.items()),
                                                       columns=['Category', 'Customer Names'])
                displayCustomers_result = displayCustomers_result.style.hide(axis="index")
                displayCustomers_result = style_commodity_customers(displayCustomers_result)

                st.write(displayCustomers_result.to_html(), unsafe_allow_html=True)

    with st.expander(f"Customer Invoicing Comparison - Raw SWMS Invoicing Data for 12 Months", expanded=True):

        ############################### Side for Uploading excel File ###############################

        date_holder, start_date_range, end_date_range, country_label, refreshButton, date_holder2 = st.columns(
            (4, 1, 1, 1, 1, 4))

        invoice_rates["Country"] = ["Australia" if " AU - " in str(x) else "New Zealand" for x in
                                    invoice_rates["Cost_Center"]]

        start_date = invoice_rates.formatted_date.dropna()

        default_startDate = start_date.min()
        default_endDate = start_date.max()

        startDate = start_date_range.date_input("Start Date", default_startDate, format="YYYY-MM-DD", disabled=True)
        endDate = end_date_range.date_input("End Date", default_endDate, format="YYYY-MM-DD", disabled=True)
        # countries = ["Australia", "New Zealand"]
        countries = invoice_rates["Country"].unique()
        country_selected = country_label.selectbox("Country", countries, index=0)

        refreshButton.text("")
        refreshButton.text("")

        if refreshButton.button("Refresh", type="secondary", icon="🔃"):
            invoice_rates = invoice_rates[
                (invoice_rates.InvoiceDate > str(startDate)) & (invoice_rates.InvoiceDate <= str(endDate))]
        else:
            invoice_rates = invoice_rates

        invoice_rates = invoice_rates[invoice_rates.Country == country_selected]

        invoice_rates["Calumo Description"] = invoice_rates.apply(sub_category_classification, axis=1)

        invoice_vols_by_Customer, invoice_rates_by_service = st.tabs(["📊 Customer Invoicing Detail",
                                                                      "🥇 Multi Customers Rates Per Service View",
                                                                      ])

        all_workday_Customer_names = invoice_rates.WorkdayCustomer_Name.unique()

        with(invoice_vols_by_Customer):

            select_Option1, select_Option2 = st.columns(2)
            cost_centres = invoice_rates.Cost_Center.unique()

            with select_Option1:
                sel1, sel2, sel4b_service, sel3 = st.columns((3, 3, 2, 1))

                selected_cost_centre_sel1 = sel1.selectbox("Cost Centre :", cost_centres, index=0, key=1140)
                selected_workday_customers = invoice_rates[
                    invoice_rates.Cost_Center == selected_cost_centre_sel1].sort_values(
                    by="WorkdayCustomer_Name").WorkdayCustomer_Name.unique()
                select_CC_data = invoice_rates[invoice_rates.Cost_Center == selected_cost_centre_sel1]

                display_rate = sel4b_service.selectbox("Avg Rate | UnitPrice : ", ["Avg Rate", "UnitPrice"],
                                                       index=0, key=1019)

                with sel2:
                    if selected_cost_centre_sel1:
                        select_customer = st.multiselect("Select Customer : ", selected_workday_customers,
                                                         selected_workday_customers[0], key=552)
                    else:
                        select_customer = st.multiselect("Select Customer : ", all_workday_Customer_names,
                                                         all_workday_Customer_names[0], key=554)

                Workday_Sales_Item_Name = invoice_rates.Workday_Sales_Item_Name.unique()
                selected_customer = select_CC_data[select_CC_data.WorkdayCustomer_Name.isin(select_customer)]

                Revenue_Category = selected_customer.Revenue_Category.dropna().unique()

                if display_rate == "UnitPrice":
                    selected_customer_pivot = pd.pivot_table(selected_customer,
                                                             values=["Quantity", "LineAmount"],
                                                             index=["Revenue_Category", "UnitOfMeasure",
                                                                    "UnitPrice"],
                                                             aggfunc="sum").reset_index()
                else:

                    selected_customer_pivot = pd.pivot_table(selected_customer,
                                                             values=["Quantity", "LineAmount"],
                                                             index=["Revenue_Category", "UnitOfMeasure", ],
                                                             aggfunc="sum").reset_index()

                    selected_customer_pivot[
                        "Avg Rate"] = selected_customer_pivot.LineAmount / selected_customer_pivot.Quantity

                selected_customer_pivot.Revenue_Category = selected_customer_pivot.Revenue_Category.apply(
                    proper_case)

                selected_customer_pivot = selected_customer_pivot.sort_values(by=["Revenue_Category", display_rate],
                                                                              ascending=False)

                col1, col2 = st.columns((2, 0.1))

                with col1:

                    customer_avg_plts = \
                        selected_customer[selected_customer['Revenue_Category'] == 'Storage - Renewal'][
                            "Quantity"].mean()
                    pallets_handled_in = \
                        selected_customer[selected_customer['Revenue_Category'] == 'Handling - Initial'][
                            "Quantity"].sum()

                    pallets_handled_out = \
                        selected_customer[selected_customer['Revenue_Category'] == 'Handling Out'][
                            "Quantity"].sum()

                    estimate_turn = ((pallets_handled_in + pallets_handled_in) / 2) / customer_avg_plts

                    selected_customer_pivot_table = selected_customer_pivot.style.format({
                        "LineAmount": "${0:,.2f}",
                        "Quantity": "{0:,.0f}",
                        "UnitPrice": "${0:,.2f}",
                        "Avg Rate": "${0:,.2f}",
                        "estimate_turn": "{0:,.1f}"
                    })

                    selected_customer_pivot_table = selected_customer_pivot_table.map(highlight_negative_values)

                    selected_customer_pivot_table = style_dataframe((selected_customer_pivot_table))
                    selected_customer_pivot_table = selected_customer_pivot_table.hide(axis="index")

                    st.write(selected_customer_pivot_table.to_html(), unsafe_allow_html=True)

                    # Convert DataFrame to Excel
                    output606 = io.BytesIO()
                    with pd.ExcelWriter(output606) as writer:
                        selected_customer_pivot.to_excel(writer)

                    # Create a download button
                    st.download_button(
                        label="Download Excel  ⤵️",
                        data=output606,
                        file_name='invoiced_data.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=616
                    )

                skip1, skip2, skip3, skip4 = st.columns(4)

                with skip1:
                    st.metric(label=f" Customer Avg Recurring Storage", value=f"{customer_avg_plts:,.0f} Plts",
                              delta=f"{estimate_turn:,.1f} : Estimate Turn",
                              border=False)

                with skip2:
                    pallet_billed_services = ['Blast Freezing', 'Storage - Initial', 'Storage - Renewal']

                    filtered_data = \
                        selected_customer[selected_customer["Revenue_Category"].isin(pallet_billed_services)][
                            "LineAmount"].sum()
                    ancillary_data = selected_customer["LineAmount"].sum() - filtered_data

                    filtered_data_contribution = filtered_data / (filtered_data + ancillary_data)

                    st.metric(label="Rent & Storage & Blast", value=f"${filtered_data:,.0f} ",
                              delta=f"{filtered_data_contribution:,.0%} : of Revenue", border=False)

                with skip3:
                    st.metric(label="Services", value=f"${ancillary_data:,.0f}",
                              delta=f"{1 - filtered_data_contribution:,.0%} : of Revenue", border=False)

            with select_Option2:

                sel1b, sel2b, sel4b_service, sel3b = st.columns((3, 3, 2, 1))

                selected_cost_centre_b = sel1b.selectbox("Cost Centre :", cost_centres, index=0, key=635)
                selected_workday_customers_b = invoice_rates[
                    invoice_rates.Cost_Center == selected_cost_centre_b].sort_values(
                    by="WorkdayCustomer_Name").WorkdayCustomer_Name.unique()
                select_CC_data_b = invoice_rates[invoice_rates.Cost_Center == selected_cost_centre_b]

                display_rate = sel4b_service.selectbox("Avg Rate | UnitPrice : ", ["Avg Rate", "UnitPrice"],
                                                       index=0, key=1129)

                with sel2b:
                    if selected_cost_centre_b:
                        select_customer_b = st.multiselect("Select Customer : ", selected_workday_customers_b,
                                                           selected_workday_customers_b[0], key=642)
                    else:
                        select_customer_b = st.multiselect("Select Customer : ", all_workday_Customer_names,
                                                           all_workday_Customer_names[0], key=644)

                Workday_Sales_Item_Name_b = invoice_rates.Workday_Sales_Item_Name.unique()
                selected_customer_b = select_CC_data_b[
                    select_CC_data_b.WorkdayCustomer_Name.isin(select_customer_b)]

                Revenue_Category_b = selected_customer_b.Revenue_Category.dropna().unique()

                if display_rate == "UnitPrice":
                    selected_customer_pivot_b = pd.pivot_table(selected_customer_b,
                                                               values=["Quantity", "LineAmount"],
                                                               index=["Revenue_Category", "UnitOfMeasure",
                                                                      "UnitPrice"],
                                                               aggfunc="sum").reset_index()
                else:
                    selected_customer_pivot_b = pd.pivot_table(selected_customer_b,
                                                               values=["Quantity", "LineAmount"],
                                                               index=["Revenue_Category", "UnitOfMeasure"],
                                                               aggfunc="sum").reset_index()

                selected_customer_pivot_b[
                    "Avg Rate"] = selected_customer_pivot_b.LineAmount / selected_customer_pivot_b.Quantity

                selected_customer_pivot_b.Revenue_Category = selected_customer_pivot_b.Revenue_Category.apply(
                    proper_case)

                selected_customer_pivot_b = selected_customer_pivot_b.sort_values(
                    by=["Revenue_Category", display_rate],
                    ascending=False)

                col1b, col2b = st.columns((2, 0.1))

                with col1b:

                    customer_avg_plts_b = \
                        selected_customer_b[selected_customer_b['Revenue_Category'] == 'Storage - Renewal'][
                            "Quantity"].mean()
                    pallets_handled_in_b = \
                        selected_customer_b[selected_customer_b['Revenue_Category'] == 'Handling - Initial'][
                            "Quantity"].sum()

                    pallets_handled_out_b = \
                        selected_customer_b[selected_customer_b['Revenue_Category'] == 'Handling Out'][
                            "Quantity"].sum()

                    estimate_turn_b = ((pallets_handled_in_b + pallets_handled_in_b) / 2) / customer_avg_plts_b

                    selected_customer_pivot_table_b = selected_customer_pivot_b.style.format({
                        "LineAmount": "${0:,.2f}",
                        "Quantity": "{0:,.0f}",
                        "Avg Rate": "${0:,.2f}",
                        "UnitPrice": "${0:,.2f}",
                    })

                    selected_customer_pivot_table_b = selected_customer_pivot_table_b.map(highlight_negative_values)

                    selected_customer_pivot_table_b = style_dataframe((selected_customer_pivot_table_b))
                    selected_customer_pivot_table_b = selected_customer_pivot_table_b.hide(axis="index")

                    st.write(selected_customer_pivot_table_b.to_html(), unsafe_allow_html=True)

                    # Convert DataFrame to Excel
                    output681 = io.BytesIO()
                    with pd.ExcelWriter(output681) as writer:
                        selected_customer_pivot.to_excel(writer)

                    # Create a download button
                    st.download_button(
                        label="Download Excel  ⤵️",
                        data=output681,
                        file_name='invoiced_data.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=691
                    )

                skip1b, skip2b, skip3b, skip4b = st.columns(4)

                with skip1b:
                    # delta_value = (customer_avg_plts / site_avg_plts_occupied)/100
                    # formartted_value = "{:.2%}".format(delta_value)
                    st.metric(label=f" Customer Avg Recurring Storage", value=f"{customer_avg_plts_b:,.0f} Plts",
                              delta=f"~{estimate_turn_b:,.1f}: Estimate Turn",
                              border=False)

                with skip2b:
                    pallet_billed_services_b = ['Blast Freezing', 'Storage - Initial', 'Storage - Renewal']

                    filtered_data_b = \
                        selected_customer_b[selected_customer_b["Revenue_Category"].isin(pallet_billed_services_b)][
                            "LineAmount"].sum()
                    ancillary_data_b = selected_customer_b["LineAmount"].sum() - filtered_data_b

                    filtered_data_contribution_b = filtered_data_b / (filtered_data_b + ancillary_data_b)

                    st.metric(label="Rent & Storage & Blast", value=f"${filtered_data_b:,.0f} ",
                              delta=f"{filtered_data_contribution_b:,.0%} : of Revenue", border=False)

                with skip3b:
                    st.metric(label="Services", value=f"${ancillary_data_b:,.0f}",
                              delta=f"{1 - filtered_data_contribution_b:,.0%} : of Revenue", border=False)

            st.text("")
            st.text("")

        with(invoice_rates_by_service):

            select_Option1_service, select_Option2_service = st.columns((1.5, 0.5))

            with select_Option1_service:
                sel1_service, sel2_service, sel3_service, sel4b_service = st.columns((2, 2, 2, 1))

                selected_cost_centre_sel1 = sel1_service.selectbox("Cost Centre :", cost_centres, index=0, key=1115)

                selected_workday_customers = invoice_rates[
                    invoice_rates.Cost_Center == selected_cost_centre_sel1].sort_values(
                    by="WorkdayCustomer_Name").WorkdayCustomer_Name.unique()

                select_CC_data = invoice_rates[invoice_rates.Cost_Center == selected_cost_centre_sel1]

                all_workday_Services = select_CC_data.Revenue_Category.unique()

                with sel2_service:
                    if selected_cost_centre_sel1:
                        select_service = st.selectbox("Service Charge : ", all_workday_Services,
                                                      index=0, key=1124)
                    else:
                        select_service = st.multiselect("Service Charge : ", all_workday_Services,
                                                        all_workday_Services[0], key=1127)

                all_unit_of_measures = select_CC_data[
                    select_CC_data.Revenue_Category == select_service].UnitOfMeasure.unique()

                with sel3_service:
                    select_UOM = st.multiselect("Unit of Measure : ", all_unit_of_measures,
                                                all_unit_of_measures, key=1407)

                Workday_Sales_Item_Name = invoice_rates.Workday_Sales_Item_Name.unique()

                selected_customer = select_CC_data[select_CC_data.Revenue_Category == select_service]

                selected_customer = selected_customer[selected_customer.UnitOfMeasure.isin(select_UOM)]

                Revenue_Category = selected_customer.Revenue_Category.dropna().unique()

                display_rate = sel4b_service.selectbox("Avg Rate | UnitPrice : ", ["Avg Rate", "UnitPrice"],
                                                       index=0, key=1152)

                if display_rate == "UnitPrice":
                    selected_customer_pivot = pd.pivot_table(selected_customer,
                                                             values=["Quantity", "LineAmount"],
                                                             index=["WorkdayCustomer_Name", "UnitOfMeasure",
                                                                    "UnitPrice"],
                                                             aggfunc="sum").reset_index()
                else:

                    selected_customer_pivot = pd.pivot_table(selected_customer,
                                                             values=["Quantity", "LineAmount"],
                                                             index=["WorkdayCustomer_Name", "UnitOfMeasure", ],
                                                             aggfunc="sum").reset_index()

                    selected_customer_pivot[
                        "Avg Rate"] = selected_customer_pivot.LineAmount / selected_customer_pivot.Quantity

                selected_customer_pivot.WorkdayCustomer_Name = selected_customer_pivot.WorkdayCustomer_Name.apply(
                    proper_case)

                selected_customer_pivot = selected_customer_pivot.sort_values(by="WorkdayCustomer_Name",
                                                                              ascending=True)

                col1_service, col2_service = st.columns((2, 0.1))

                with col1_service:

                    customer_avg_plts = \
                        selected_customer[selected_customer['Revenue_Category'] == 'Storage - Renewal'][
                            "Quantity"].sum()
                    pallets_handled_in = \
                        selected_customer[selected_customer['Revenue_Category'] == 'Handling - Initial'][
                            "Quantity"].sum()

                    pallets_handled_out = \
                        selected_customer[selected_customer['Revenue_Category'] == 'Handling Out'][
                            "Quantity"].sum()

                    estimate_turn = ((pallets_handled_in + pallets_handled_in) / 2) / customer_avg_plts

                    selected_customer_pivot_table = selected_customer_pivot.style.format({
                        "LineAmount": "${0:,.2f}",
                        "Quantity": "{0:,.0f}",
                        "UnitPrice": "${0:,.2f}",
                        "Avg Rate": "${0:,.2f}",
                        "estimate_turn": "{0:,.1f}"
                    })

                    selected_customer_pivot_table = selected_customer_pivot_table.map(highlight_negative_values)

                    selected_customer_pivot_table = style_dataframe(selected_customer_pivot_table)
                    selected_customer_pivot_table = selected_customer_pivot_table.hide(axis="index")

                    st.write(selected_customer_pivot_table.to_html(), unsafe_allow_html=True)

                    # Convert DataFrame to Excel
                    output1171 = io.BytesIO()
                    with pd.ExcelWriter(output1171) as writer:
                        selected_customer_pivot.to_excel(writer)

                    # Create a download button
                    st.download_button(
                        label="Download Excel  ⤵️",
                        data=output1171,
                        file_name='invoiced_data.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key=1181
                    )

                skip1_service, skip2_service, skip3_service, skip4_service = st.columns(4)

                with skip1_service:
                    pass
                    # st.metric(label=f" Customer Avg Recurring Storage", value=f"{customer_avg_plts:,.0f} Plts",
                    #           delta=f"{estimate_turn:,.1f} : Estimate Turn",
                    #           border=False)

                with skip2_service:
                    pallet_billed_services = ['Blast Freezing', 'Storage - Initial', 'Storage - Renewal']

                    filtered_data = \
                        selected_customer[selected_customer["Revenue_Category"].isin(pallet_billed_services)][
                            "LineAmount"].sum()
                    ancillary_data = selected_customer["LineAmount"].sum() - filtered_data

                    filtered_data_contribution = filtered_data / (filtered_data + ancillary_data)

                    st.metric(label="Rent & Storage & Blast", value=f"${filtered_data:,.0f} ", border=False)

                with skip3_service:
                    st.metric(label="Services", value=f"${ancillary_data:,.0f}", border=False)

            with select_Option2_service:

                st.subheader("2025 Service Rates Comparison", divider='rainbow')

                def1_service, wt1_service, def2_service = st.columns((1, 3, 1))

                with wt1_service:
                    st.markdown(f"""
                                \n
                            __Raw Invoice Data :__ \n
                            Note - data used excludes Credit Note entries. \n
                            \n
                            1. Select __Service Charges__ and respective __Unit of Measure.__ \n
                            2. Multiple Customer entries may mean in year Rate Review. \n
                            3. Some Sites have billing inconsistencies (e.g bill by amount).\n 
                            4. When Avg Rate Selected - may not be the exact Customer Rate as per Rate Card. 
                            
                        """)

    st.divider()

    st.subheader("Review of Customer Rates at Site", divider="rainbow")

    #     ###############################  Box Plots Over and under Indexed Customers  ###############################

    v1, v2, v3, plot_UOM = st.columns((0.5, 1, 0.15, 1))

    region_list_rates = sorted(customer_rate_cards.Region.unique())
    region_list_rates = np.insert(region_list_rates, 0, "All")
    selected_region_filter = v1.selectbox('Select Region', region_list_rates, index=1)
    selected_region_rate_cards = customer_rate_cards[customer_rate_cards["Region"] == selected_region_filter]

    site_list_rates = sorted(selected_region_rate_cards.Site.unique())
    selected_site_rate_cards = v2.multiselect("Select Site :", site_list_rates, site_list_rates, key=1433)

    # site_list_rates = region_list_rates.Site.unique()
    plots_unit_of_measure = sorted(selected_region_rate_cards.Prop.unique())

    display_box_customers = v3.selectbox("Select View :", ["All", "Outliers Only"], index=0)
    selected_Plot_UOM = plot_UOM.multiselect("Unit of Measure : ", plots_unit_of_measure,
                                             plots_unit_of_measure, key=1426)

    selected_region_rate_cards = (customer_rate_cards
                                  if selected_region_filter == 'All'
                                  else selected_region_rate_cards[
        selected_region_rate_cards.Site.isin(selected_site_rate_cards)])

    selected_region_rate_cards = (customer_rate_cards
                                  if selected_region_filter == 'All'
                                  else selected_region_rate_cards[
        selected_region_rate_cards.Prop.isin(selected_Plot_UOM)])

    st.text("")
    st.text("")

    p1_Storage, p2_Handling, = st.columns((1, 1))

    df_pStorage = selected_region_rate_cards[selected_region_rate_cards.Description == "Storage"]

    try:

        with p1_Storage:
            df_pStorage = selected_region_rate_cards[selected_region_rate_cards.Description == "Storage"]
            df_pStorage_customers_count = len(df_pStorage)
            customer_names = df_pStorage.Customer

            q3_pStorage = np.percentile(df_pStorage.Rate.values, 75)
            q1_pStorage = np.percentile(df_pStorage.Rate.values, 25)

            # #### Storage Box plot ####################################################################
            # # Add points with Labels
            y = df_pStorage.Rate.values
            x = np.random.normal(1, 0.250, size=(len(y)))

            positive_separator = 0.1
            negative_separator = -0.2
            y_axes_offset = 0
            seperator = 0

            # Create Box Plot Figure:
            box_x_value = np.random.normal(1, 0.200, size=len(y))

            # Create a Boxplot
            fig = go.Figure()

            fig.add_trace(go.Box(
                y=y,
                boxpoints='all',
                jitter=0.8,
                pointpos=0,
                name="Storage Rate",
                line={"color": '#4D93D9'},
                marker={"size": 5, "color": 'red'},

            ))

            # Add annotations for each data point. to recheck
            for i, (val, customer) in enumerate(zip(y, customer_names)):
                # if len(customer.title()) > 9:
                #     customerName = f'{customer.title()[:9] + '...'}',
                # else:
                #     customerName = f'{customer.title()[:9] + '...'}'

                if display_box_customers == "Outliers Only":
                    if float(y[i]) < q1_pStorage or float(y[i]) > q3_pStorage:
                        seperator = negative_separator if seperator == positive_separator else positive_separator
                        fig.add_annotation(
                            x=float(x[i]),
                            y=float(y[i]),
                            showarrow=False,
                            xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            text=f'{customer.title()[:9] + '...'}',
                            font={"size": 13,
                                  "color": 'black'})

                        fig.update_layout(
                            title=dict(
                                text=f'{selected_region_filter} : Customers ({df_pStorage_customers_count}) : Storage Rates'),
                            margin=dict(l=0, r=10, b=10, t=30),
                        )
                else:
                    seperator = negative_separator if seperator == positive_separator else positive_separator
                    fig.add_annotation(
                        x=float(x[i]),
                        y=float(y[i]),
                        showarrow=False,
                        xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                        hovertext=f'{customer}, Rate:  {val}',
                        text=f'{customer.title()[:7] + '...'}',
                        font={"size": 13,
                              "color": 'black'})

                    fig.update_layout(
                        title=dict(
                            text=f'{selected_region_filter} : Customers ({df_pStorage_customers_count}) : Storage Rates'),
                        margin=dict(l=0, r=10, b=10, t=30),
                    )

            p1_Storage.plotly_chart(fig)

        # Handling Box plot ####################################################################
        with p2_Handling:
            df_pHandling = selected_region_rate_cards[selected_region_rate_cards.Description == "Pallet Handling"]
            df_pHandling_customers_count = len(df_pHandling)
            handling_customer_names = df_pHandling.Customer

            q3_pHandling = np.percentile(df_pHandling.Rate.values, 75)
            q1_pHandling = np.percentile(df_pHandling.Rate.values, 25)

            # Add points with Labels
            y = df_pHandling.Rate.values
            x = np.random.normal(1, 0.250, size=(len(y)))

            # Create Box Plot Figure:
            box_x_value = np.random.normal(1, 0.250, size=len(y))

            positive_separator = 0.1
            negative_separator = -0.2
            y_axes_offset = 0
            seperator = 0

            # Create a Boxplot
            fig = go.Figure()

            fig.add_trace(go.Box(
                y=y,
                boxpoints='all',
                jitter=0.8,
                pointpos=0,
                name="Handling Rates",
                line={"color": '#4D93D9'},
                marker={"size": 5, "color": 'red'},
            ))

            # Add annotations for each data point
            for i, (val, customer) in enumerate(zip(y, handling_customer_names)):

                if display_box_customers == "Outliers Only":
                    if float(y[i]) < q1_pHandling or float(y[i]) > q3_pHandling:
                        seperator = negative_separator if seperator == positive_separator else positive_separator
                        fig.add_annotation(
                            x=float(x[i]),
                            y=float(y[i]),
                            showarrow=False,
                            xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            text=f'{customer.title()[:7] + '...'}',
                            font={"size": 13,
                                  "color": 'black'}
                        )

                        fig.update_layout(
                            title=dict(
                                text=f'{selected_region_filter} : Customers ({df_pHandling_customers_count}): Handling Rates'),
                            margin=dict(l=0, r=10, b=10, t=30)
                        )
                else:
                    shortName = [customer.title()[:9] + '...' if len(customer.title()) > 10 else customer.title() for
                                 cus in customer]
                    seperator = negative_separator if seperator == positive_separator else positive_separator
                    fig.add_annotation(
                        x=float(x[i]),
                        y=float(y[i]),
                        showarrow=False,
                        xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                        hovertext=f'{customer}, Rate:  {val}',
                        text=f'{customer.title()[:7] + '...'}',
                        font={"size": 13,
                              "color": 'black'})

                    fig.update_layout(
                        title=dict(
                            text=f'{selected_region_filter} : Customers ({df_pHandling_customers_count}): Handling Rates'),
                        margin=dict(l=0, r=10, b=10, t=30)
                    )

            p2_Handling.plotly_chart(fig)

        st.divider()
        st.markdown("##### Carton Picking and Shrink Wrap Rates")
        st.divider()

        p3_Wrapping, p4_Cartons, = st.columns((1, 1))

        # Shrink Wrap Box plot ######################################################

        with p3_Wrapping:
            df_pWrapping = selected_region_rate_cards[selected_region_rate_cards.Description == "Shrink Wrap"]
            df_pWrapping_customers_count = len(df_pWrapping)
            wrapping_customer_names = df_pWrapping.Customer

            q3_pWrapping = np.percentile(df_pWrapping.Rate.values, 75)
            q1_pWrapping = np.percentile(df_pWrapping.Rate.values, 25)

            # Add points with Labels
            y = df_pWrapping.Rate.values
            x = np.random.normal(1, 0.250, size=(len(y)))

            positive_separator = 0.1
            negative_separator = -0.2
            y_axes_offset = 0
            seperator = 0

            # Create Box Plot Figure:
            box_x_value = np.random.normal(1, 0.250, size=len(y))

            # Create a Boxplot
            fig = go.Figure()

            fig.add_trace(go.Box(
                y=y,
                boxpoints='all',
                jitter=0.8,
                pointpos=0,
                name="Shrink Wrap Rates",
                line={"color": '#4D93D9'},
                marker={"size": 5, "color": 'red'},
            ))

            # Add annotations for each data point
            for i, (val, customer) in enumerate(zip(y, wrapping_customer_names)):

                if display_box_customers == "Outliers Only":
                    if float(y[i]) < q1_pWrapping or float(y[i]) > q3_pWrapping:
                        seperator = negative_separator if seperator == positive_separator else positive_separator
                        fig.add_annotation(
                            x=float(x[i]),
                            y=float(y[i]),
                            showarrow=False,
                            xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            text=f'{customer.title()[:7] + '...'}',
                            font={"size": 13,
                                  "color": 'black'}
                        )

                        fig.update_layout(
                            title=dict(
                                text=f'{selected_region_filter} : Customers ({df_pWrapping_customers_count}) : Shrink Wrapping Rates'),
                            margin=dict(l=0, r=10, b=10, t=20))
                else:
                    seperator = negative_separator if seperator == positive_separator else positive_separator
                    fig.add_annotation(
                        x=float(x[i]),
                        y=float(y[i]),
                        showarrow=False,
                        xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                        hovertext=f'{customer}, Rate:  {val}',
                        text=f'{customer.title()[:7] + '...'}',
                        font={"size": 13,
                              "color": 'black'})

                    fig.update_layout(
                        title=dict(
                            text=f'{selected_region_filter} : Customers ({df_pWrapping_customers_count}) : Shrink Wrapping Rates'),
                        margin=dict(l=0, r=10, b=10, t=20))

            p3_Wrapping.plotly_chart(fig)

        with p4_Cartons:
            ############################ Shrink Wrap Box plot ######################################################
            df_pCarton = selected_region_rate_cards[selected_region_rate_cards.Description == "Carton Picking"]
            df_pCarton_customers_count = len(df_pCarton)
            picking_customer_names = df_pCarton.Customer

            q3_pCarton = np.percentile(df_pCarton.Rate.values, 75)
            q1_pCarton = np.percentile(df_pCarton.Rate.values, 25)

            # Add points with Labels
            y = df_pCarton.Rate.values
            x = np.random.normal(1, 0.250, size=(len(y)))

            positive_separator = 0.1
            negative_separator = -0.2
            y_axes_offset = 0
            seperator = 0

            # Create Box Plot Figure:
            box_x_value = np.random.normal(1, 0.250, size=len(y))

            # Create a Boxplot
            fig = go.Figure()

            fig.add_trace(go.Box(
                y=y,
                boxpoints='all',
                jitter=0.8,
                pointpos=0,
                name="Carton Picking Rates",
                line={"color": '#4D93D9'},
                marker={"size": 5, "color": 'red'},
            ))

            # Add annotations for each data point
            for i, (val, customer) in enumerate(zip(y, picking_customer_names)):

                if display_box_customers == "Outliers Only":
                    if float(y[i]) < q1_pCarton or float(y[i]) > q3_pCarton:
                        seperator = negative_separator if seperator == positive_separator else positive_separator
                        fig.add_annotation(
                            x=float(x[i]),
                            y=float(y[i]),
                            showarrow=False,
                            xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            text=f'{customer.title()[:7] + '...'}',
                            font={"size": 13,
                                  "color": 'black'}
                        )

                        fig.update_layout(
                            title=dict(
                                text=f'{selected_region_filter} : Customers ({df_pCarton_customers_count}) : Carton Picking Rates'),
                            margin=dict(l=0, r=10, b=10, t=20))
                else:
                    seperator = negative_separator if seperator == positive_separator else positive_separator
                    fig.add_annotation(
                        x=float(x[i]),
                        y=float(y[i]),
                        showarrow=False,
                        xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                        hovertext=f'{customer}, Rate:  {val}',
                        text=f'{customer.title()[:7] + '...'}',
                        font={"size": 13,
                              "color": 'black'}
                    )

                    fig.update_layout(
                        title=dict(
                            text=f'{selected_region_filter} : Customers ({df_pCarton_customers_count}) : Carton Picking Rates'),
                        margin=dict(l=0, r=10, b=10, t=20))

            p4_Cartons.plotly_chart(fig)

    except Exception as e:
        st.markdown(f"##############  Data pulled for this Customer or Site does not have all the Columns needed "
                    f"to display info ")

    # st.subheader("", divider="rainbow")

    st.text("")
    st.text("")

    with st.expander(f"Extended View", expanded=False):

        st.subheader("Other Customer Rates Analysis Invoiced at Site ", divider="rainbow")

        _, weeks_input, _ = st.columns((4, 1, 4))

        weeks_ago = weeks_input.number_input(f"Last 4 wks ending {endDate} ", value=4, key=1782)

        changed_weeks_ago_date = date_three_weeks_ago(endDate, number_of_weeks=weeks_ago)

        default_3_weeks_ago_startDate = datetime.date(changed_weeks_ago_date.year, changed_weeks_ago_date.month,
                                                      changed_weeks_ago_date.day)

        invoice_rates.Cost_Center = invoice_rates.Cost_Center.apply(extract_site)

        site_list_rates = invoice_rates.Cost_Center.unique()

        box_whisker_invoice_rates = invoice_rates[invoice_rates.InvoiceDate > str(default_3_weeks_ago_startDate)]

        v1, service_view, v2 = st.columns(3)

        last_cost_center = len(site_list_rates) - 1

        selected_site_rate_cards = v1.selectbox("Site Selection for Rate Cards :", site_list_rates,
                                                index=0, key=1816)

        st.text("")
        st.text("")

        box_whisker_invoice_rates = box_whisker_invoice_rates[
            box_whisker_invoice_rates.Cost_Center == selected_site_rate_cards]

        all_workday_Services = box_whisker_invoice_rates.Revenue_Category.unique()

        last_service = len(all_workday_Services) - 1

        selected_service_view = service_view.selectbox("Site Selection for Rate Cards :", all_workday_Services,
                                                       index=last_service,
                                                       key=1778)

        box_whisker_invoice_rates = box_whisker_invoice_rates[
            box_whisker_invoice_rates.Revenue_Category == selected_service_view]

        box_whisker_invoice_rates = pd.pivot_table(box_whisker_invoice_rates,
                                                   values=["Quantity", "LineAmount"],
                                                   index=["Revenue_Category", "UnitOfMeasure", "WorkdayCustomer_Name",
                                                          "UnitPrice"],
                                                   aggfunc="sum").reset_index()

        # box_whisker_invoice_rates[
        #     "Avg Rate"] = box_whisker_invoice_rates.LineAmount / box_whisker_invoice_rates.Quantity

        box_whisker_invoice_rates[
            "Avg Rate"] = box_whisker_invoice_rates.UnitPrice

        p5_Others, _, p5_Table, = st.columns((2, 0.5, 2))

        with p5_Others:

            try:

                display_box_customers = st.selectbox("Select Customers to View :", ["All", "Outliers Only"], index=0,
                                                     key=1761)
                st.text("")
                st.text("")

                ############################ Shrink Wrap Box plot ######################################################

                df_pServiceView = box_whisker_invoice_rates[
                    box_whisker_invoice_rates.Revenue_Category == selected_service_view]

                all_workday_Services = df_pServiceView.Revenue_Category.unique()

                all_unit_of_measures = df_pServiceView.UnitOfMeasure.unique()

                select_UOM = v2.multiselect("Unit of Measure : ", all_unit_of_measures,
                                            all_unit_of_measures, key=1906)

                df_pServiceView = df_pServiceView[df_pServiceView.UnitOfMeasure.isin(select_UOM)]

                df_pCarton_customers_count = len(df_pServiceView)
                picking_customer_names = df_pServiceView.WorkdayCustomer_Name

                q3_pCarton = np.percentile(df_pServiceView["Avg Rate"].values, 75)
                q1_pCarton = np.percentile(df_pServiceView["Avg Rate"].values, 25)

                # Add points with Labels
                y = df_pServiceView["Avg Rate"].values
                x = np.random.normal(1, 0.250, size=(len(y)))

                positive_separator = 0.1
                negative_separator = -0.2
                y_axes_offset = 0
                seperator = 0

                # Create Box Plot Figure:
                box_x_value = np.random.normal(1, 0.250, size=len(y))

                # Create a Boxplot
                fig = go.Figure()

                fig.add_trace(go.Box(
                    y=y,
                    boxpoints='all',
                    jitter=0.8,
                    pointpos=0,
                    name=f"{selected_service_view} Rates",
                    line={"color": '#4D93D9'},
                    marker={"size": 5, "color": 'red'},
                ))

                # Add annotations for each data point
                for i, (val, customer) in enumerate(zip(y, picking_customer_names)):
                    shortName = [customer.title()[:9] + '...' if len(customer.title()) > 10 else customer.title() for
                                 cus in customer]
                    if display_box_customers == "Outliers Only":
                        if float(y[i]) < q1_pCarton or float(y[i]) > q3_pCarton:
                            seperator = negative_separator if seperator == positive_separator else positive_separator
                            fig.add_annotation(
                                x=float(x[i]),
                                y=float(y[i]),
                                showarrow=False,
                                xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                                hovertext=f'{customer.title()}, Rate:  {val}',
                                text=f"{shortName[i]} + 5 ",
                                font={"size": 13,
                                      "color": 'black'}
                            )

                            fig.update_layout(
                                title=dict(
                                    text=f'{selected_site_rate_cards} : Customers ({df_pCarton_customers_count}) : {selected_service_view} Rates'),
                                margin=dict(l=0, r=10, b=10, t=20))
                    else:
                        seperator = negative_separator if seperator == positive_separator else positive_separator
                        fig.add_annotation(
                            x=float(x[i]),
                            y=float(y[i]),
                            showarrow=False,
                            xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer.title()}, Rate:  {val}',
                            text=f"{shortName[i]}",
                            font={"size": 13,
                                  "color": 'black'}
                        )

                        fig.update_layout(
                            title=dict(
                                text=f'{selected_site_rate_cards} : Customers ({df_pCarton_customers_count}) : {selected_service_view} Rates'),
                            margin=dict(l=0, r=10, b=10, t=20))

                p5_Others.plotly_chart(fig)

            except Exception as e:
                st.markdown('#### No Data ')

        with p5_Table:

            try:

                selected_customer = df_pServiceView

                display_rate = "UnitPrice"
                if display_rate == "UnitPrice":
                    selected_customer_pivot = pd.pivot_table(selected_customer,
                                                             values=["Quantity", "LineAmount"],
                                                             index=["WorkdayCustomer_Name", "UnitOfMeasure",
                                                                    "UnitPrice"],
                                                             aggfunc="sum").reset_index()
                else:

                    selected_customer_pivot = pd.pivot_table(selected_customer,
                                                             values=["Quantity", "LineAmount"],
                                                             index=["WorkdayCustomer_Name", "UnitOfMeasure", ],
                                                             aggfunc="sum").reset_index()

                    selected_customer_pivot[
                        "Avg Rate"] = selected_customer_pivot.LineAmount / selected_customer_pivot.Quantity

                selected_customer_pivot.WorkdayCustomer_Name = selected_customer_pivot.WorkdayCustomer_Name.apply(
                    proper_case)

                selected_customer_pivot = selected_customer_pivot.sort_values(by="WorkdayCustomer_Name",
                                                                              ascending=True)

                selected_customer_pivot_table = selected_customer_pivot.style.format({
                    "LineAmount": "${0:,.2f}",
                    "Quantity": "{0:,.0f}",
                    "UnitPrice": "${0:,.2f}",
                    "Avg Rate": "${0:,.2f}",
                    "estimate_turn": "{0:,.1f}"
                })

                selected_customer_pivot_table = selected_customer_pivot_table.map(highlight_negative_values)

                selected_customer_pivot_table = style_dataframe(selected_customer_pivot_table)
                selected_customer_pivot_table = selected_customer_pivot_table.hide(axis="index")

                st.write(selected_customer_pivot_table.to_html(), unsafe_allow_html=True)
            except Exception as e:
                st.markdown('#### No Data ')
    st.subheader("", divider="rainbow")

    with AboutTab:
        st.markdown('''##### Arthur Rusike:
                                    *Report any Code breaks or when you see Red Error Message*
                                '''
                    )
