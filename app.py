import io
import random
from io import BytesIO
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import streamlit as st
from helper_functions import style_dataframe, load_profitbility_Summary_model, load_rates_standardisation, \
    load_specific_xls_sheet
from streamlit import session_state as ss
from plotly.subplots import make_subplots

from site_detailed_view_helper_functions import load_invoices_model, profitability_model

st.set_page_config(page_title="Project Renaissance Analysis", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")


def extract_cc(cost_centre):
    return cost_centre[:9]


def extract_site(value):
    return value.split(" - ")[1].lower().title()


def extract_name(value):
    return value.split(" [")[0].lower().title()


def extract_cost_name(value):
    return value.split(" AU - ")[1].lower().title()


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
            # Main Dashboard
            invoice_rates = load_invoices_model(uploaded_invoicing_data)
            st.divider()

            st.markdown("#### Dates Covered by Invoice File")
            start_date = invoice_rates.formatted_date.dropna()
            st.markdown(f'###### Report Start Date  : {start_date.min()}')
            st.markdown(f'###### Report End Date  : {start_date.max()}')
            st.divider()


    except Exception as e:

        st.subheader('To use with this WebApp 📊  - Upload below : \n '
                     '1. Profitability Summary File \n'
                     '2. Customer Rates Summary File \n'
                     '3. Customer Invoicing Data')

# Main Dashboard ##############################

if uploaded_file and customer_rates_file and uploaded_invoicing_data:

    DetailsTab, SiteTab, AboutTab = st.tabs(["📊 Quick View", "🥇 Site Detailed View",
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
    site_list = list(profitability_summary_file.Site.unique())

    #  Details Tab ###############################

    with (DetailsTab):
        space_holder_1, title_holder, cost_centres_selection, space_holder3 = st.columns(4)

        title_holder.subheader("Project Renaissance", divider="blue")
        selected_site = cost_centres_selection.multiselect("Site :", site_list, site_list[0])

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

        with st.expander("Full Years Comparison - GAAP (USD) ", expanded=True):

            section1, section2, section3 = st.columns((
                                                                   len(selected_site) * 0.20,
                                                                   len(selected_site) * 1,
                                                                   len(selected_site) * 1))

            with section1:
                selected_years = []

                t1, t2, t3, t4, = st.columns((4, 1, 1, 1))

                with t1:
                    st.text("")
                    st.text("")
                    t1.markdown("##### Years")
                    _2025 = t1.checkbox('2025', True)
                    _2024 = t1.checkbox('2024', True)
                    _2023 = t1.checkbox('2023', True)

                    if _2025:
                        selected_years.append(2025)

                    if _2024:
                        selected_years.append(2024)

                    if _2023:
                        selected_years.append(2023)

            budget_data_2025 = budget_data_2025[budget_data_2025.Year.isin(selected_years)]
            budget_data_2025 = budget_data_2025[budget_data_2025.Site.isin(selected_site)]

            budget_data_2025 = budget_data_2025.groupby(by=["Site", "Year"]).sum().reset_index()
            budget_data_2025 = budget_data_2025.sort_values(by=["Site", "Year"], ascending=[False, False])

            budget_data_2025_pivot = budget_data_2025
            budget_data_2025.insert(0, "Description",
                                    budget_data_2025_pivot["Site"] + " : " + budget_data_2025_pivot["Year"].astype(
                                        str))
            budget_data_2025_pivot = budget_data_2025_pivot.set_index("Description")


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


            budget_data_2025_pivot["Revenue"] = budget_data_2025_pivot["Revenue"].apply(format_for_currency)
            budget_data_2025_pivot["Ebitda"] = budget_data_2025_pivot["Ebitda"].apply(format_for_currency)
            budget_data_2025_pivot["Rev Per Pallet"] = budget_data_2025_pivot["Rev Per Pallet"].apply(
                format_for_float_currency)
            budget_data_2025_pivot["Ebitda Per Pallet"] = budget_data_2025_pivot["Ebitda Per Pallet"].apply(
                format_for_float_currency)
            budget_data_2025_pivot["Economic Utilization"] = budget_data_2025_pivot["Economic Utilization"].apply(
                format_for_percentage)
            budget_data_2025_pivot["Economic OHP"] = (budget_data_2025_pivot["Economic OHP"] / 12).apply(
                format_for_int)
            budget_data_2025_pivot["Labour to Tot. Rev"] = budget_data_2025_pivot["Labour to Tot. Rev"].apply(
                format_for_percentage)
            budget_data_2025_pivot["DL to Svcs Rev"] = budget_data_2025_pivot["DL to Svcs Rev"].apply(
                format_for_percentage)
            budget_data_2025_pivot["Direct Labor / Hour"] = budget_data_2025_pivot["Direct Labor / Hour"].apply(
                format_for_float_currency)
            budget_data_2025_pivot["Turn"] = budget_data_2025_pivot["Turn"].apply(format_for_float)
            budget_data_2025_pivot["Rev - Storage & Blast"] = budget_data_2025_pivot["Rev - Storage & Blast"].apply(
                format_for_percentage)
            budget_data_2025_pivot["Rev - Services"] = budget_data_2025_pivot["Rev - Services"].apply(
                format_for_percentage)

            budget_data_2025_pivot = budget_data_2025_pivot.T

            budget_data_2025_pivot = budget_data_2025_pivot[2:]

            styles = [
                {
                    'selector': ' tr:hover',
                    'props': [
                        ('border', '1px solid #4CAF50'),
                        ('background-color', 'wheat'),

                    ]
                },
                {
                    'selector': ' tr',
                    'props': [
                        ('text-align', 'right'),
                        ('font-size', '12px'),
                        ('font-family', 'sans-serif, Arial'),

                    ]
                }

            ]

            for column in budget_data_2025_pivot.columns:
                styles.append({

                    'selector': f'th.col{budget_data_2025_pivot.columns.get_loc(column)}',
                    'props': [
                        ('background-color', '#305496'),
                        ('width', 'auto'),
                        ('color', 'white'),
                        ('font-family', 'sans-serif, Arial'),
                        ('font-size', '12px'),
                        ('text-align', 'center'),
                        ('border', '2px solid white')
                    ],

                }
                )

            budget_data_2025_pivot = budget_data_2025_pivot.style.set_table_styles(styles)

            section2.markdown("##### Site Financial Summary - USD")

            section2.write(budget_data_2025_pivot.to_html(), unsafe_allow_html=True)

            with section3:

                bar1,_bar_dash, bar2 = st.columns((2,0.5,2))

                budget_data_2025["Reve Per Pallet"] = budget_data_2025["Rev Per Pallet"] - budget_data_2025["Ebitda Per Pallet"]

                def format_for_float_2_currency(value):
                    value = value/1000_000
                    return "${0:,.2f}M".format(value)


                fig = px.bar(budget_data_2025, x='Year',
                             y='Revenue',
                             title=f"Site Revenue",
                             text=budget_data_2025["Revenue"].map(format_for_float_2_currency)
                             )
                fig.update_traces(width=0.7,
                                  marker=dict(cornerradius=30),
                                  )
                bar1.plotly_chart(fig,key=303 , use_container_width=True)

                fig = px.bar(budget_data_2025, x='Year',
                             y='Ebitda',
                             title="Site Ebitda",
                             text=budget_data_2025["Ebitda"].map(format_for_float_2_currency)

                             )
                fig.update_traces(width=0.5,
                                  marker=dict(cornerradius=30),
                                  )
                bar2.plotly_chart(fig,key=308,  use_container_width=True )


            # with section3:
            #     st.text("")
            #     st.text("")
            #     st.text("")
            #
            #     kpi1, kpi2, = st.columns(2)
            #
            #     # with kpi1:
            #     revenue_2024 = select_Site_data[" Revenue"].sum() / 1000
            #     revenue_2023 = select_Site_data_2023_profitability[" Revenue"].sum() / 1000
            #     delta = f"{((revenue_2024 / revenue_2023) - 1):,.1%}"
            #
            #     kpi1.metric(label=f"FY24: Revenue - 000s", value=f"{revenue_2024:,.0f}", delta=f"{delta} vs LY ")
            #
            #     # with kpi2:
            #     ebitda_2024 = select_Site_data["EBITDA $"].sum() / 1000
            #     ebitda_2023 = select_Site_data_2023_profitability["EBITDA $"].sum() / 1000
            #     delta = f"{((ebitda_2024 / ebitda_2023) - 1):,.1%}"
            #
            #     ebitda_margin = f"{(ebitda_2024 / revenue_2024):,.1%}"
            #
            #     kpi2.metric(label=f"FY24: EBITDA - 000s", value=f"{ebitda_2024:,.0f}",
            #                 delta=f"{delta} | FY24 → {ebitda_margin}   ")
            #     st.text("")
            #     st.text("")
            #
            #     kpi3, kpi4, = st.columns(2)
            #
            #     dominic_report_2024 = dominic_report[
            #         (dominic_report.Year == 2024) & (dominic_report.Month == "December")]
            #     dominic_report_2024["Site"] = dominic_report_2024["Cost Centers"].apply(extract_site)
            #     dominic_report_2024 = dominic_report_2024[dominic_report_2024.Site.isin(selected_site)]
            #
            #     dominic_report_2023 = dominic_report[
            #         (dominic_report.Year == 2023) & (dominic_report.Month == "December")]
            #     dominic_report_2023["Site"] = dominic_report_2023["Cost Centers"].apply(extract_site)
            #     dominic_report_2023 = dominic_report_2023[dominic_report_2023.Site.isin(selected_site)]
            #
            #     economic_pallets_2024 = dominic_report_2024["Economic Pallet Total"].sum()
            #     capacity_2024 = dominic_report_2024.Denominator.sum()
            #
            #     economic_pallets_2023 = dominic_report_2023["Economic Pallet Total"].sum()
            #     capacity_2023 = dominic_report_2023.Denominator.sum()
            #
            #     dec_2024_occupancy = economic_pallets_2024 / capacity_2024
            #     dec_2023_occupancy = economic_pallets_2023 / capacity_2023
            #
            #     delta = f"{((dec_2024_occupancy / dec_2023_occupancy) - 1):,.1%}"
            #
            #     kpi3.metric(label=f"Economic Occupancy", value=f"{dec_2024_occupancy:,.1%}", delta=f"{delta} ")
            #
            #     # with kpi4:
            #     volume_guarantee_2024 = select_Site_data_2023_pallets["VG Pallets - Dec-2024"].sum()
            #     volume_guarantee_2023 = select_Site_data_2023_pallets["VG Pallets - Dec-2023"].sum()
            #
            #     delta = f"{(volume_guarantee_2024 / capacity_2024):,.1%} : of Capacity is VG "
            #
            #     kpi4.metric(label=f"Volume Guarantee", value=f"{volume_guarantee_2024:,.0f}", delta=f"{delta} ")
            st.text("")
            st.text("")

        st.subheader("", divider='rainbow')

        st.text("")
        st.text("")

        with st.expander("2024 Financial and KPI Summaries - Adjusted for Market Rent", expanded=False):

            st.text("")
            st.text("")
            st.text("")

            _k1, kpi1, kpi2, _k_2 = st.columns(4)

            # with kpi1:
            revenue_2024 = select_Site_data[" Revenue"].sum() / 1000
            revenue_2023 = select_Site_data_2023_profitability[" Revenue"].sum() / 1000
            delta = f"{((revenue_2024 / revenue_2023) - 1):,.1%}"

            kpi1.metric(label=f"FY24: Revenue - 000s", value=f"{revenue_2024:,.0f}", delta=f"{delta} vs LY ")

            # with kpi2:
            ebitda_2024 = select_Site_data["EBITDA $"].sum() / 1000
            ebitda_2023 = select_Site_data_2023_profitability["EBITDA $"].sum() / 1000
            delta = f"{((ebitda_2024 / ebitda_2023) - 1):,.1%}"

            ebitda_margin = f"{(ebitda_2024 / revenue_2024):,.1%}"

            kpi2.metric(label=f"FY24: EBITDA - 000s", value=f"{ebitda_2024:,.0f}",
                        delta=f"{delta} | FY24 → {ebitda_margin}   ")
            st.text("")
            st.text("")

            _k_3, kpi3, kpi4, _k_4 = st.columns(4)

            dominic_report_2024 = dominic_report[
                (dominic_report.Year == 2024) & (dominic_report.Month == "December")]
            dominic_report_2024["Site"] = dominic_report_2024["Cost Centers"].apply(extract_site)
            dominic_report_2024 = dominic_report_2024[dominic_report_2024.Site.isin(selected_site)]

            dominic_report_2023 = dominic_report[
                (dominic_report.Year == 2023) & (dominic_report.Month == "December")]
            dominic_report_2023["Site"] = dominic_report_2023["Cost Centers"].apply(extract_site)
            dominic_report_2023 = dominic_report_2023[dominic_report_2023.Site.isin(selected_site)]

            economic_pallets_2024 = dominic_report_2024["Economic Pallet Total"].sum()
            capacity_2024 = dominic_report_2024.Denominator.sum()

            economic_pallets_2023 = dominic_report_2023["Economic Pallet Total"].sum()
            capacity_2023 = dominic_report_2023.Denominator.sum()

            dec_2024_occupancy = economic_pallets_2024 / capacity_2024
            dec_2023_occupancy = economic_pallets_2023 / capacity_2023

            delta = f"{((dec_2024_occupancy / dec_2023_occupancy) - 1):,.1%}"

            kpi3.metric(label=f"Economic Occupancy", value=f"{dec_2024_occupancy:,.1%}", delta=f"{delta} ")

            # with kpi4:
            volume_guarantee_2024 = select_Site_data_2023_pallets["VG Pallets - Dec-2024"].sum()
            volume_guarantee_2023 = select_Site_data_2023_pallets["VG Pallets - Dec-2023"].sum()

            delta = f"{(volume_guarantee_2024 / capacity_2024):,.1%} : of Capacity is VG "

            kpi4.metric(label=f"Volume Guarantee", value=f"{volume_guarantee_2024:,.0f}", delta=f"{delta} ")
            st.text("")
            st.text("")

        fin1, fin2 = st.columns(2)

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
        key_metrics_data_pivot.insert(3, "Ebitdar - %", key_metrics_data_pivot['EBITDAR $'] / key_metrics_data_pivot[' Revenue'])


        key_metrics_data_pivot_Hume["Ebitda - %"] = key_metrics_data_pivot_Hume['EBITDA $'] / \
                                                    key_metrics_data_pivot_Hume[' Revenue']

        key_metrics_data_pivot_Hume.insert(2,"Ebitdar - %", key_metrics_data_pivot_Hume['EBITDAR $'] / \
                                           key_metrics_data_pivot_Hume[' Revenue']  )

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
            fin1.write(key_metrics_data_pivot.to_html(), unsafe_allow_html=True)
            if "Hume Rd" == selected_site[0]:
                st.text("")
                st.text("")
                st.text("")
                fin1.write(key_metrics_data_pivot_Hume.to_html(), unsafe_allow_html=True)

        key_metrics_data_Hume_KPI =  key_metrics_data.groupby("Site").sum()

        # key_metrics_data["Capacity"] =   key_metrics_data["Pal Cap"] / len(key_metrics_data[key_metrics_data == key_metrics_data["CC"] ])

        test = key_metrics_data[["CC","Pal Cap"]].groupby("CC").mean()
        key_metrics_data = key_metrics_data.groupby("CC").sum()

        DL_Handling = key_metrics_data["Direct Labor Expense,\n$"]
        Ttl_Labour = key_metrics_data["Total Labor Expense, $"]
        Service_Revenue = key_metrics_data["Service Rev (Handling+ Case Pick+ Other Rev)"]

        key_metrics_data["Capacity"] =  key_metrics_data["Pallet"]
        key_metrics_data["Pallets"] = key_metrics_data["Pallet"]
        key_metrics_data["Occ-%"] = key_metrics_data["Pallet"] / key_metrics_data["Capacity"]
        key_metrics_data["Rev | Plt"] = key_metrics_data[" Revenue"] / (key_metrics_data["Pallet"] * 52)
        key_metrics_data["EBITDA | Plt"] = key_metrics_data["EBITDA $"] / (key_metrics_data["Pallet"] * 52)
        key_metrics_data["Rent | Plt"] = (key_metrics_data["Rent Expense,\n$"] / key_metrics_data["Pallet"] )/52
        key_metrics_data["Turn"] = ((key_metrics_data["TTP p.w."] / 2) * 52) / key_metrics_data["Pallet"]
        key_metrics_data["DL %"] = DL_Handling / Service_Revenue
        key_metrics_data["LTR %"] = Ttl_Labour / key_metrics_data[" Revenue"]
        key_metrics_data["Rev | SQM"] = key_metrics_data[" Revenue"] / key_metrics_data["sqm"]
        key_metrics_data["EBITDA | SQM"] = key_metrics_data["EBITDA $"] / key_metrics_data["sqm"]
        key_metrics_data["Rent | SQM"] = key_metrics_data['Rent Expense,\n$'] / key_metrics_data['sqm']


        # key_metrics_data["Site Pal Cap psqm"] = key_metrics_data['Site Pal Cap psqm']

        kpi_key_metrics_data_pivot = key_metrics_data[
            [   "Pallet",
                "Capacity",
            "Pallets", "Occ-%", "Rev | Plt","Rent | Plt", "EBITDA | Plt", "Turn", "DL %", "LTR %", "Rev | SQM","Rent | SQM",
             "EBITDA | SQM",

             ]]

        kpi_key_metrics_data_pivot = pd.concat([kpi_key_metrics_data_pivot, test], axis=1, )
        kpi_key_metrics_data_pivot["Capacity"] = kpi_key_metrics_data_pivot["Pal Cap"]
        kpi_key_metrics_data_pivot["Occ-%"] = kpi_key_metrics_data_pivot["Pallet"] / kpi_key_metrics_data_pivot["Capacity"]


        budget_data_2025_bench_mark = kpi_key_metrics_data_pivot

        kpi_key_metrics_data_pivot = kpi_key_metrics_data_pivot.style.hide(axis="index")
        kpi_key_metrics_data_pivot = kpi_key_metrics_data_pivot.hide(["Pallet","Pal Cap"], axis="columns")


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
            "Rent | Plt":'${0:,.2f}',
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
            key_metrics_data_Hume_KPI["Occ-%"] = key_metrics_data_Hume_KPI["Pallet"] / key_metrics_data_Hume_KPI["Capacity"]
            key_metrics_data_Hume_KPI["Rev | Plt"] = key_metrics_data_Hume_KPI[" Revenue"] / (key_metrics_data_Hume_KPI["Pallet"] * 52)
            key_metrics_data_Hume_KPI["EBITDA | Plt"] = key_metrics_data_Hume_KPI["EBITDA $"] / (key_metrics_data_Hume_KPI["Pallet"] * 52)
            key_metrics_data_Hume_KPI["Rent | Plt"] = (key_metrics_data_Hume_KPI["Rent Expense,\n$"] / key_metrics_data_Hume_KPI["Pallet"])/52
            key_metrics_data_Hume_KPI["Turn"] = ((key_metrics_data_Hume_KPI["TTP p.w."] / 2) * 52) / key_metrics_data_Hume_KPI["Pallet"]
            key_metrics_data_Hume_KPI["DL %"] = DL_Handling / Service_Revenue
            key_metrics_data_Hume_KPI["LTR %"] = Ttl_Labour / key_metrics_data_Hume_KPI[" Revenue"]
            key_metrics_data_Hume_KPI["Rev | SQM"] = key_metrics_data_Hume_KPI[" Revenue"] / key_metrics_data_Hume_KPI["sqm"]
            key_metrics_data_Hume_KPI["EBITDA | SQM"] = key_metrics_data_Hume_KPI["EBITDA $"] / key_metrics_data_Hume_KPI["sqm"]
            key_metrics_data_Hume_KPI["Rent | SQM"] = key_metrics_data_Hume_KPI['Rent Expense,\n$'] / key_metrics_data_Hume_KPI['sqm']
            # key_metrics_data["Site Pal Cap psqm"] = key_metrics_data['Site Pal Cap psqm']

            kpi_key_metrics_data_pivot_Hume_KPI = key_metrics_data_Hume_KPI[
                [
                    "Capacity",
                    "Pallets", "Occ-%", "Rev | Plt","Rent | Plt", "EBITDA | Plt", "Turn", "DL %", "LTR %", "Rev | SQM",
                    "EBITDA | SQM",

                    # "Rent | SQM"
                    # "Site Pal Cap psqm"
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
                "Rent | Plt":'${0:,.2f}',
            })

            kpi_key_metrics_data_pivot_Hume_KPI = kpi_key_metrics_data_pivot_Hume_KPI.map(highlight_negative_values)

            kpi_key_metrics_data_pivot_Hume_KPI = style_dataframe(kpi_key_metrics_data_pivot_Hume_KPI)
            st.text("")
            st.text("")
            st.text("")
            if "Hume Rd" == selected_site[0]:
                fin2.write(kpi_key_metrics_data_pivot_Hume_KPI.to_html(), unsafe_allow_html=True)


        st.divider()

        # TOP 10 Customers ##################################################

        with st.expander("Top Customers Review"):

            profitability_by_Customer, customer_network_profitability =  st.tabs(["📊 Top Site Customers",
                                                                                  "🥇 Multi-Site Customer Profitability View",
                                                                                  ])

            with(profitability_by_Customer):
                st.subheader("Top Customers", divider='rainbow')

                st.text("")

                s1, s2, s3, s4, s5 = st.columns(5)

                with s1:
                    # s1_a, s1_b = st.columns(2)
                    selected_cost_centre = st.selectbox("Select Cost Centre :", cost_centres, index=0, )

                selected_CC = selected_cost_centre[:9]
                select_CC_data = profitability_summary_file[profitability_summary_file["CC"] == selected_CC]
                select_CC_data_2023_profitability = profitability_2023[
                    profitability_2023["CC"] == selected_CC]
                select_CC_data_2023_pallets = customer_pallets[customer_pallets["CC"] == selected_CC]

                select_CC_data = select_CC_data.query("Customer != '.All Other [.All Other]'")
                select_CC_data = select_CC_data.query("Pallet > 1")
                select_CC_data["Rank"] = select_CC_data[" Revenue"].rank(method='max')
                select_CC_data["Rev psqm"] = select_CC_data[" Revenue"] / select_CC_data["sqm"]
                treemap_data = select_CC_data.filter(
                    ['Name', " Revenue", 'EBITDAR $', 'EBITDAR Margin\n%' ,'EBITDA $', 'EBITDA Margin\n%','Revenue per Pallet', 'EBITDA per Pallet',
                     'Pallet', 'TTP p.w.', 'Rank', 'Turn', "DLH% \n(DL / Service Rev)", "LTR % (Labour to Rev %)", 'Rev psqm',
                     'EBITDA psqm', 'sqm', 'Rent psqm', 'Storage Revenue, $', 'Blast Freezing Revenue, $','RSB Rev per OHP',
                     'Service Rev per TPP', 'Rent Expense,\n$'
                     ])

                # treemap_data.rename(columns={'EBITDAR Margin\n%': 'EBITDAR %','EBITDA Margin\n%':'EBITDA %','EBITDA per Pallet':'EBITDA | Plt',
                #                              'Revenue per Pallet': 'Rev | Plt','DLH% \n(DL / Service Rev)': 'DL Ratio' })


                treemap_data.insert(7,"Rent | Plt", (treemap_data['Rent Expense,\n$'] / (treemap_data.Pallet * 52)))

                display_data = treemap_data
                size = len(display_data)
                customer_view_size = s2.number_input("Filter Bottom Outlier Customers", value=size)
                display_data = display_data.query(f"Rank > {size - customer_view_size} ")
                display_data_pie = display_data
                rank_display_data = display_data
                score_carding_data = display_data
                display_data["RSB"] = (display_data['Storage Revenue, $'] + display_data[
                    'Blast Freezing Revenue, $']) / display_data[" Revenue"]
                display_data["Services"] = 1 - display_data["RSB"]
                display_data.rename(columns={"DLH% \n(DL / Service Rev)": "DL ratio", "LTR % (Labour to Rev %)": "LTR - %",
                                             'EBITDAR Margin\n%': 'EBITDAR %','EBITDA Margin\n%':'EBITDA %','EBITDA per Pallet':'EBITDA | Plt',
                                             'Revenue per Pallet': 'Rev | Plt'}, inplace=True)

                treemap_graph_data = display_data
                display_data = display_data.style.hide(axis="index")

                display_data = display_data.hide(['Rank', 'sqm', 'Storage Revenue, $', 'Blast Freezing Revenue, $','Rent psqm',
                                                  'Rent Expense,\n$',
                                                  ],
                                                 axis="columns")

                display_data = style_dataframe(display_data)

                display_data = display_data.format({
                    " Revenue": '${0:,.0f}',
                    'EBITDAR $':"${0:,.0f}",
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
                    "Rent | Plt":"${0:,.2f}",
                    'Rent psqm': "${0:,.2f}",
                    'Turn': "{0:,.2f}",
                    'RSB': "{0:,.2%}",
                    'Services': "{0:,.2%}",
                    'RSB Rev per OHP':"${0:,.2f}",
                    'Service Rev per TPP':"${0:,.2f}",

                })

                display_data = display_data.map(highlight_negative_values)
                display_data.background_gradient(subset=['EBITDA %'], cmap="RdYlGn")

                # Create a download button
                with s3:
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

                st.write(display_data.to_html(), unsafe_allow_html=True, use_container_width=True)

                st.text("")
                st.text("")
                st.text("")

            with(customer_network_profitability):
                st.subheader("Multi-Site Customer Profitability", divider='rainbow')

                st.text("")
                st.text("")

                s1, s2, s3 = st.columns(3)

                with s1:
                    profitability_summary_file = profitability_summary_file.query("Customer != '.All Other [.All Other]'")

                    multi_site_customers = profitability_summary_file.sort_values(by="Name").Name.unique()
                    default_customers = [multi_site_customers[8], multi_site_customers[27],multi_site_customers[28] ]

                    select_network_customers = st.multiselect("Select Customers : ", multi_site_customers,
                                                              default_customers, key=792)

                select_network_data = profitability_summary_file[profitability_summary_file.Name.isin(select_network_customers)]

                select_network_data = select_network_data.query("Pallet > 1")
                select_network_data["Rank"] = select_network_data[" Revenue"].rank(method='max')
                select_network_data["Rev psqm"] = select_network_data[" Revenue"] / select_network_data["sqm"]

                treemap_network_data = select_network_data.filter(
                    ['Name',"Site","Cost Center", " Revenue", 'EBITDAR $', 'EBITDAR Margin\n%' ,'EBITDA $', 'EBITDA Margin\n%','Revenue per Pallet', 'EBITDA per Pallet',
                     'Pallet', 'TTP p.w.', 'Rank', 'Turn', "DLH% \n(DL / Service Rev)", "LTR % (Labour to Rev %)", 'Rev psqm',
                     'EBITDA psqm', 'sqm', 'Rent psqm', 'Storage Revenue, $', 'Blast Freezing Revenue, $','RSB Rev per OHP',
                     'Service Rev per TPP', 'Rent Expense,\n$','Avg. Storage Rate (Storage $/Pallet)','Avg. Handling Rate (Handling $/TTP)'
                     ])
                treemap_network_data["Cost Center"] = [ x.split(" - ")[1] +" - "+ x.split(" - ")[2] for x in  treemap_network_data["Cost Center"]   ]

                treemap_network_data.insert(7,"Rent | Plt", (treemap_network_data['Rent Expense,\n$'] / (treemap_network_data.Pallet * 52)))
                treemap_network_data = treemap_network_data.sort_values(by=" Revenue", ascending=False)

                display_network_data = treemap_network_data
                size = len(display_network_data)

                network_customer_view_size = s2.number_input("Filter Bottom Outlier Customers", value=size)
                display_network_data = display_network_data.query(f"Rank > {size - network_customer_view_size} ")



                display_network_data["RSB"] = (display_network_data['Storage Revenue, $'] + display_network_data[
                    'Blast Freezing Revenue, $']) / display_network_data[" Revenue"]
                display_network_data["Services"] = 1 - display_network_data["RSB"]
                display_network_data.rename(columns={"DLH% \n(DL / Service Rev)": "DL ratio", "LTR % (Labour to Rev %)": "LTR - %",
                                             'EBITDAR Margin\n%': 'EBITDAR %','EBITDA Margin\n%':'EBITDA %','EBITDA per Pallet':'EBITDA | Plt',
                                             'Revenue per Pallet': 'Rev | Plt', 'Avg. Storage Rate (Storage $/Pallet)':'Avg. Storage Rate' ,
                                             'Avg. Handling Rate (Handling $/TTP)': 'Avg. Handling Rate'

                                            }, inplace=True)

                # treemap_graph_data = display_data
                display_network_data = display_network_data.style.hide(axis="index")

                display_network_data = display_network_data.hide(['Rank', 'sqm', 'Storage Revenue, $', 'Blast Freezing Revenue, $','Rent psqm',
                                                  'Rent Expense,\n$', 'Rev psqm', 'EBITDA psqm','Avg. Storage Rate',
                                                                  'Avg. Handling Rate'
                                                  ],
                                                 axis="columns")

                display_network_data = style_dataframe(display_network_data)

                display_network_data = display_network_data.format({
                    " Revenue": '${0:,.0f}',
                    'EBITDAR $':"${0:,.0f}",
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
                    "Rent | Plt":"${0:,.2f}",
                    'Rent psqm': "${0:,.2f}",
                    'Turn': "{0:,.2f}",
                    'RSB': "{0:,.2%}",
                    'Services': "{0:,.2%}",
                    'RSB Rev per OHP':"${0:,.2f}",
                    'Service Rev per TPP':"${0:,.2f}",

                })

                display_network_data = display_network_data.map(highlight_negative_values)
                display_network_data.background_gradient(subset=['EBITDA %'], cmap="RdYlGn")

                # Create a download button
                with s3:
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

                st.write(display_network_data.to_html(), unsafe_allow_html=True, use_container_width=True)

                st.text("")
                st.text("")
                st.text("")


        with st.expander("Customer Invoicing Comparison - Raw SWMS Invoicing Data for 2024", expanded=False):

            ############################### Side for Uploading excel File ###############################

            all_workday_Customer_names = invoice_rates.WorkdayCustomer_Name.unique()

            select_Option1, select_Option2 = st.columns(2)

            with select_Option1:
                sel1, sel2, sel3 = st.columns((3, 3, 1))

                selected_cost_centre_sel1 = sel1.selectbox("Cost Centre :", cost_centres, index=0, key=544)
                selected_workday_customers = invoice_rates[
                    invoice_rates.Cost_Center == selected_cost_centre_sel1].sort_values(
                    by="WorkdayCustomer_Name").WorkdayCustomer_Name.unique()
                select_CC_data = invoice_rates[invoice_rates.Cost_Center == selected_cost_centre_sel1]

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

                selected_customer_pivot = pd.pivot_table(selected_customer,
                                                         values=["Quantity", "LineAmount"],
                                                         index=["Revenue_Category", "Workday_Sales_Item_Name"],
                                                         aggfunc="sum").reset_index()
                selected_customer_pivot[
                    "Avg Rate"] = selected_customer_pivot.LineAmount / selected_customer_pivot.Quantity

                col1, col2 = st.columns((2, 0.1))

                with col1:

                    customer_avg_plts = selected_customer[selected_customer['Revenue_Category'] == 'Storage - Renewal'][
                        "Quantity"].mean()
                    pallets_handled_in = \
                        selected_customer[selected_customer['Revenue_Category'] == 'Handling - Initial'][
                            "Quantity"].sum()

                    pallets_handled_out = selected_customer[selected_customer['Revenue_Category'] == 'Handling Out'][
                        "Quantity"].sum()

                    estimate_turn = ((pallets_handled_in + pallets_handled_in) / 2) / customer_avg_plts

                    selected_customer_pivot_table = selected_customer_pivot.style.format({
                        "LineAmount": "${0:,.2f}",
                        "Quantity": "{0:,.0f}",
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

                sel1b, sel2b, sel3b = st.columns((3, 3, 1))

                selected_cost_centre_b = sel1b.selectbox("Cost Centre :", cost_centres, index=0, key=635)
                selected_workday_customers_b = invoice_rates[
                    invoice_rates.Cost_Center == selected_cost_centre_b].sort_values(
                    by="WorkdayCustomer_Name").WorkdayCustomer_Name.unique()
                select_CC_data_b = invoice_rates[invoice_rates.Cost_Center == selected_cost_centre_b]

                with sel2b:
                    if selected_cost_centre_b:
                        select_customer_b = st.multiselect("Select Customer : ", selected_workday_customers_b,
                                                           selected_workday_customers_b[0], key=642)
                    else:
                        select_customer_b = st.multiselect("Select Customer : ", all_workday_Customer_names,
                                                           all_workday_Customer_names[0], key=644)

                Workday_Sales_Item_Name_b = invoice_rates.Workday_Sales_Item_Name.unique()
                selected_customer_b = select_CC_data_b[select_CC_data_b.WorkdayCustomer_Name.isin(select_customer_b)]

                Revenue_Category_b = selected_customer_b.Revenue_Category.dropna().unique()

                selected_customer_pivot_b = pd.pivot_table(selected_customer_b,
                                                           values=["Quantity", "LineAmount"],
                                                           index=["Revenue_Category", "Workday_Sales_Item_Name"],
                                                           aggfunc="sum").reset_index()
                selected_customer_pivot_b[
                    "Avg Rate"] = selected_customer_pivot_b.LineAmount / selected_customer_pivot_b.Quantity

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

        with st.expander("Select Customer to Benchmark", expanded=True):
            customer_list = rank_display_data.Name.unique()
            st.subheader("Customer KPIs vs Site")
            _bench1, _bench2, _bench3, = st.columns(3)
            selected_customer = _bench2.selectbox("Select Customer", customer_list,
                                                  index=0, placeholder="Customer to Benchmark")

            rank_display_data_benchmark = rank_display_data.loc[rank_display_data["Name"] == selected_customer]

            site_benchmark = selected_cost_centre.split(" - ")[1].strip().title()

            selected_site.append(str(site_benchmark))

            site_benchmark_cc = selected_cost_centre.split(" AU")[0]

            budget_data_2025_bench_mark = budget_data_2025_bench_mark.loc[
                budget_data_2025_bench_mark.index == site_benchmark_cc]

            _a, b1, b2, b3, _b = st.columns((1, 3, 3, 3, 1))

            with b1:
                rank_display_data_benchmark["Site Rev Per Pallet"] = budget_data_2025_bench_mark["Rev | Plt"].values

                df = rank_display_data_benchmark[["Rev | Plt", "Site Rev Per Pallet"]]

                # st.dataframe(rank_display_data_benchmark)

                df = df.T
                fig = make_subplots(rows=2, cols=2, shared_yaxes=False, column_widths=[100, 100],
                                    row_heights=[250, 250],
                                    horizontal_spacing=1, vertical_spacing=0, shared_xaxes=True
                                    )

                fig.add_trace(
                    go.Bar(x=df.loc["Rev | Plt"], y=[f'{selected_customer}'], name="Rev | Pallet",
                           orientation='h',
                           text=df.loc["Rev | Plt"].map(format_for_float_currency),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30), showlegend=False
                           ), 1, 1
                )

                fig.add_trace(
                    go.Bar(x=df.loc["Site Rev Per Pallet"], y=[f'{site_benchmark}'], name="Site Rev Per Pallet",
                           orientation='h',
                           text=df.loc["Site Rev Per Pallet"].map(format_for_float_currency),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30, color="#156082"), showlegend=False
                           ), 2, 1,
                )

                fig.update_layout(height=250, width=600, title_text="Revenue Per Pallet vs Site",
                                  )

                b1.plotly_chart(fig, use_container_width=True)

            with b2:
                rank_display_data_benchmark["Site Ebitda Per Pallet"] = budget_data_2025_bench_mark[
                    "EBITDA | Plt"].values

                df = rank_display_data_benchmark[["EBITDA | Plt", "Site Ebitda Per Pallet"]]
                df = df.T
                fig = make_subplots(rows=2, cols=2, shared_yaxes=False, column_widths=[100, 100],
                                    row_heights=[250, 250],
                                    horizontal_spacing=1, vertical_spacing=0, shared_xaxes=True
                                    )

                fig.add_trace(
                    go.Bar(x=df.loc["EBITDA | Plt"], y=[f'{selected_customer}'], name="Ebitda Per Pallet",
                           orientation='h',
                           text=df.loc["EBITDA | Plt"].map(format_for_float_currency),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30), showlegend=False
                           ), 1, 1
                )

                fig.add_trace(
                    go.Bar(x=df.loc["Site Ebitda Per Pallet"], y=[f'{site_benchmark}'], name="Site Ebitda Per Pallet",
                           orientation='h',
                           text=df.loc["Site Ebitda Per Pallet"].map(format_for_float_currency),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30, color="#156082"), showlegend=False
                           ), 2, 1,
                )

                fig.update_layout(height=250, width=600, title_text="EBITDA Per Pallet vs Site")

                b2.plotly_chart(fig, use_container_width=True, key=455)

            with b3:
                rank_display_data_benchmark["Site Pallet Turns"] = budget_data_2025_bench_mark["Turn"].values

                df = rank_display_data_benchmark[["Turn", "Site Pallet Turns"]]
                df = df.T
                fig = make_subplots(rows=2, cols=2, shared_yaxes=False, column_widths=[100, 100],
                                    row_heights=[250, 250],
                                    horizontal_spacing=1, vertical_spacing=0, shared_xaxes=True
                                    )

                fig.add_trace(
                    go.Bar(x=df.loc["Turn"], y=[f'{selected_customer}'], name="Turn",
                           orientation='h',
                           text=df.loc["Turn"].map(format_for_float),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30), showlegend=False
                           ), 1, 1
                )

                fig.add_trace(
                    go.Bar(x=df.loc["Site Pallet Turns"], y=[f'{site_benchmark}'], name="Site Pallet Turn",
                           orientation='h',
                           text=df.loc["Site Pallet Turns"].map(format_for_float),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30, color="#156082"), showlegend=False
                           ), 2, 1,
                )

                fig.update_layout(height=250, width=600, title_text="Customer's Turn vs Site")

                b3.plotly_chart(fig, use_container_width=True, key=593)

            _c, c1, c2, _d = st.columns((1, 3, 3, 1))

            with c1:
                rank_display_data_benchmark["Site DL Per Pallet"] = budget_data_2025_bench_mark["DL %"].values

                df = rank_display_data_benchmark[["DL ratio", "Site DL Per Pallet"]]
                df = df.T
                fig = make_subplots(rows=2, cols=2, shared_yaxes=False, column_widths=[100, 100],
                                    row_heights=[250, 250],
                                    horizontal_spacing=1, vertical_spacing=0, shared_xaxes=True
                                    )

                fig.add_trace(
                    go.Bar(x=df.loc["DL ratio"], y=[f'{selected_customer}'], name="DL Ratio",
                           orientation='h',
                           text=df.loc["DL ratio"].map(format_for_percentage),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30), showlegend=False
                           ), 1, 1
                )

                fig.add_trace(
                    go.Bar(x=df.loc["Site DL Per Pallet"], y=[f'{site_benchmark}'], name="Site DL Per Pallet",
                           orientation='h',
                           text=df.loc["Site DL Per Pallet"].map(format_for_percentage),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30, color="#156082"), showlegend=False
                           ), 2, 1,
                )

                fig.update_layout(height=250, width=600, title_text="Customer DL Ratio vs Site")

                b1.plotly_chart(fig, use_container_width=True, key=486)

            with c2:
                rank_display_data_benchmark["Site LTR"] = budget_data_2025_bench_mark["LTR %"].values

                df = rank_display_data_benchmark[["LTR - %", "Site LTR"]]
                df = df.T
                fig = make_subplots(rows=2, cols=2, shared_yaxes=False, column_widths=[100, 100],
                                    row_heights=[250, 250],
                                    horizontal_spacing=1, vertical_spacing=0, shared_xaxes=True
                                    )

                fig.add_trace(
                    go.Bar(x=df.loc["LTR - %"], y=[f'{selected_customer}'], name="LTR", orientation='h',
                           text=df.loc["LTR - %"].map(format_for_percentage),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30), showlegend=False
                           ), 1, 1
                )

                fig.add_trace(
                    go.Bar(x=df.loc["Site LTR"], y=[f'{site_benchmark}'], name="LTR", orientation='h',
                           text=df.loc["Site LTR"].map(format_for_percentage),
                           textfont=dict(color='white'),
                           marker=dict(cornerradius=30, color="#156082"), showlegend=False
                           ), 2, 1,
                )

                fig.update_layout(height=250, width=600, title_text="Customer LTR vs Site")

                b2.plotly_chart(fig, use_container_width=True, key=512)


        with st.expander("Score Card | Customer Grading ", expanded=True):
            st.subheader("Customers Score Card", divider='rainbow')

            def1, wt1, def2 = st.columns(3)

            with wt1:
                revenue_per_pallet_weighting = st.number_input("Rev Per Pallet Weighting", value=20)
                ebitda_pallet_weighting = st.number_input("Ebitda Per Pallet Weighting", value=40)
                direct_labour_ratio_weighting = st.number_input("Direct Labour Ratio Weighting", value=20)
                stock_turn_weighting = st.number_input("Stock Turn Weighting", value=20)
                margin_weighting = st.number_input("Margin % Weighting", value=0)

            def1.markdown(f"""
                    __Score Card Definitions__ \n
                    Score - assign a Performance metric based on Customer's KPIs relative to it's Contribution at site. \n
                    
                    1. Identified KPIs and Their Weightings
                    First, identify the key performance indicators (KPIs) you want to use and assign a weight to each based on its importance. For example:
                    -	Revenue Per Pallet:...................{revenue_per_pallet_weighting}%
                    -	EBITDA Per Pallet:.....................{ebitda_pallet_weighting}%
                    -	Direct Labour Ratio:..................{direct_labour_ratio_weighting}%
                    -	Stock Turns:..............................{stock_turn_weighting}%
                    -	Margin:......................................{margin_weighting}%
                    	
                    Total Weighting :{revenue_per_pallet_weighting +
                                      +ebitda_pallet_weighting +
                                      stock_turn_weighting +
                                      direct_labour_ratio_weighting +
                                      margin_weighting}%
                                        
                    2. Normalize KPIs - (Scaling) 
                    3. Calculate the Weighted Score
                    
                """)

            def2.markdown("""
                    __Score Card Bands__ \n
                    
                    
                    |Band                  |    Call to Action                       |
                    |:---------------------|:------------------------------------|
                    |1.00 : High Risk      | Needs immediate review and action   |
                    |2.00 : Unsatisfactory | Identify areas of Improvement       |
                    |2.50 : Red Flag       | Future problem review now           |
                    |3.00 : Satisfactory   | Satisfactory to have Customer       |
                    |3.50 : Good           | Good to have Customer               |
                    |4.00 : Excellent      | Excellent to have Customer          |
                    
                    
                """)

        # st.divider()

        with st.expander("Top 10 Customers - Activity Rank", expanded=True):

            rank1, rank2, rank3 = st.columns((1,4,1))
            with rank2:

                rank_display_data["Stock Turn Times"] = (
                        ((rank_display_data["TTP p.w."] / 2) * 52) / rank_display_data["Pallet"])

                rank_display_data.insert(1,"Rev Rank",rank_display_data[" Revenue"].rank(method='max', ascending=False))
                rank_display_data.insert(2, "EBITDA", rank_display_data["EBITDA $"].rank(method='max', ascending=False))
                rank_display_data.insert(3,"Rev | Pallet", rank_display_data["Rev | Plt"].rank(method='max', ascending=False))
                rank_display_data["EBITDA | Plt"] = rank_display_data["EBITDA | Plt"].rank(method='max', ascending=False)

                rank_display_data["Margin %"] = rank_display_data["EBITDA %"]

                rank_display_data["Pallets"] = rank_display_data["Pallet"].rank(method='max', ascending=False)
                rank_display_data["TTP"] = rank_display_data["TTP p.w."].rank(method='max', ascending=False)
                rank_display_data["Turns"] = rank_display_data["Stock Turn Times"].rank(method='max', ascending=False)


                def calculate_score_card(df):

                    revenue_score = df["Rev | Plt"]
                    revenue_score_min = 0
                    revenue_score_max = df["Rev | Plt"].max()
                    normalised_revenue_score = (revenue_score - revenue_score_min) / (
                            revenue_score_max - revenue_score_min)

                    ebitda_score = df["EBITDA | Plt"]
                    ebitda_score_min = 0
                    ebitda_score_max = df["EBITDA | Plt"].max()
                    normalised_ebitda_score = (ebitda_score - ebitda_score_min) / (ebitda_score_max - ebitda_score_min)

                    dl_ratio_score = df["LTR - %"]
                    dl_ratio_score_min = 0
                    dl_ratio_score_max = df["LTR - %"].max()
                    normalised_dl_ratio_score = (dl_ratio_score - dl_ratio_score_min
                                                 ) / (dl_ratio_score_max - dl_ratio_score_min)

                    turn_score = df["Turn"]
                    turn_score_min = 0
                    turn_score_max = df["Turn"].max()
                    normalised_turn_score = (turn_score - turn_score_min
                                             ) / (turn_score_max - turn_score_min)

                    margin_score = df['EBITDA %']
                    margin_score_min = 0
                    margin_score_max = df['EBITDA %'].max()
                    normalised_pallet_score = (margin_score - margin_score_min
                                               ) / (margin_score_max - margin_score_min)

                    final_score = (normalised_revenue_score * (revenue_per_pallet_weighting / 100)) + (
                            normalised_ebitda_score * (ebitda_pallet_weighting / 100)) + (
                                          normalised_pallet_score * (margin_weighting / 100)) + (
                                          normalised_dl_ratio_score * (direct_labour_ratio_weighting / 100)) + (
                                          normalised_turn_score * (stock_turn_weighting / 100))

                    df["Score"] = (final_score * 4) + 1

                    return df


                calculate_score_card(rank_display_data)

                rank_display_data["Score"] = [1 if y < 1000 else x for x, y in
                                              zip(rank_display_data["Score"], rank_display_data["EBITDA $"])]


                def assign_score_card(score):
                    if score >= 4:
                        return "Excellent"
                    elif score >= 3.50:
                        return "Good"
                    elif score >= 3.00:
                        return "Satisfactory"
                    elif score >= 2.50:
                        return "Red Flag"
                    elif score >= 2.00:
                        return "Unsatisfactory"
                    else:
                        return "High Risk"


                rank_display_data["Score Card"] = rank_display_data["Score"].apply(assign_score_card)


                def add_comment(score_card):
                    match score_card:
                        case "High Risk":
                            return "Needs immediate review and action"
                        case "Unsatisfactory":
                            return "Identify areas of Improvement"
                        case "Red Flag":
                            return "Future problem review now"
                        case "Satisfactory":
                            return "Satisfactory to have Customer"
                        case "Good":
                            return "Good to have Customer"
                        case "Excellent":
                            return "Excellent to have Customer"


                rank_display_data["Comment"] = (rank_display_data["Score Card"].apply(add_comment))

                # columns_to_hide = [" Revenue", 'EBITDAR $' , 'EBITDA $', 'EBITDA %',
                #                    'Rev | Plt', 'sqm', 'Rent psqm', 'Storage Revenue, $',
                #                    'Blast Freezing Revenue, $','EBITDAR %', 'Rent | Plt',
                #                    'Pallet', 'TTP p.w.', 'Stock Turn Times', ]

                rank_display_data = rank_display_data.style.hide(axis="index")
                rank_display_data = rank_display_data.hide(
                    [" Revenue", 'EBITDA $', 'EBITDA %', 'Storage Revenue, $', 'Blast Freezing Revenue, $',
                     'Rev | Plt', 'Pallet', "RSB", "Services",'EBITDAR $','Rent | Plt','EBITDAR %','Turn',
                     'TTP p.w.', 'sqm', 'Rent psqm', 'Stock Turn Times', "Rev psqm", "EBITDA psqm", 'Rank', "DL ratio",
                     "LTR - %",'RSB Rev per OHP','Service Rev per TPP', 'Rent Expense,\n$'  ], axis="columns")

                rank_display_data = rank_display_data.format({
                    "Rev Rank": '{0:,.0f}',
                    "EBITDA": '{0:,.0f}',
                    "Margin %": '{0:,.2%}',
                    "Rev | Pallet": '{0:,.0f}',
                    "EBITDA | Plt": '{0:,.0f}',
                    'Turns': '{0:,.0f}',
                    "Pallets": '{0:,.0f}',
                    "TTP": '{0:,.0f}',
                    "Score": '{0:,.2f}',
                })

                rank_display_data.background_gradient(subset=['Score'], cmap="RdYlGn")
                # def assign_score_card():
                #    score =   rank_display_data["Rev Rank"] +

                rank_display_data = style_dataframe(rank_display_data)
                rank_display_data = rank_display_data.map(highlight_negative_values)

                rank2.write(rank_display_data.to_html(), unsafe_allow_html=True, use_container_width=True)

                output1a = io.BytesIO()
                with pd.ExcelWriter(output1a) as writer:
                    rank_display_data.to_excel(writer, sheet_name='export_data', index=False)

                # Create a download button

                st.text("")
                st.text("")
                st.download_button(
                    label="👆 Download ⤵️",
                    data=output1a,
                    file_name='customer_revenue.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    key=1132,
                )

        st.divider()

        with st.expander("Revenue vs Profitability Contribution - Analysis "):

            st.subheader(f'Revenue vs Profitability Contribution - Analysis : {selected_cost_centre}',
                         divider='rainbow')

            graph1, graph2 = st.columns(2)

            customers = treemap_graph_data['Name']

            color_columns = [" Revenue", 'EBITDA $',
                             'EBITDA %',
                             'Rev | Plt', 'EBITDA | Plt',
                             'Pallet',
                             'TTP p.w.']
            # remark = select_CC_data['EBITDA Margin\n%']  # Replace with EBITDA per Pallet Through Put
            # margin = select_CC_data['EBITDA Margin\n%']  # Replace with EBITDA per Pallet Through Put

            with graph1:
                graph1_size, graph1_color, space_holder_3, space_holder_4 = st.columns(4)

                box_size_view = graph1_size.selectbox("Select Option for Box Size:", color_columns, key=1142, index=3)
                box_color_view = graph1_color.selectbox("Select Option for Box Color:", color_columns, key=1143,
                                                        index=4)

                # Plotting HEAT MAP - tow Show Revenue and Profitability ############################

                fig = px.treemap(treemap_graph_data,
                                 path=[customers],
                                 values=treemap_graph_data[box_size_view],
                                 color=treemap_graph_data[box_color_view],
                                 # color_continuous_scale=['#074F69', '#0C769E', '#61CBF3','#94DCF8',"#CAEDFB"]
                                 color_continuous_scale='RdYlGn'

                                 )

                # fig.update_traces(root_color="red",)

                st.plotly_chart(fig, filename='Chart.html')

            with graph2:
                treemap_graph_data = treemap_graph_data.filter(['Name', " Revenue", 'EBITDA $', 'EBITDA %',
                                                                'Rev | Plt', 'EBITDA | Plt', 'Pallet',
                                                                'TTP p.w.'])
                # with st.expander("Filter Customer to Display"):
                #     st.data_editor(treemap_graph_data, num_rows='dynamic')

                # treemap_graph_data = treemap_graph_data[
                #     (treemap_graph_data[box_size_view] & treemap_graph_data[box_color_view] )]
                x_value = str(box_size_view)
                y_value = str(box_color_view)

                fig = px.scatter(treemap_graph_data, x=y_value,
                                 y=x_value,
                                 size=treemap_graph_data[" Revenue"],
                                 color="Name",
                                 hover_name="Name",
                                 text="Name",
                                 log_x=False,
                                 size_max=60)
                st.plotly_chart(fig)

        st.divider()

        st.subheader("Review of Customer Rates at Site", divider="rainbow")

        #     ###############################  Box Plots Over and under Indexed Customers  ###############################

        site_list_rates = customer_rate_cards.Site.unique()
        v1, v2 = st.columns(2)

        selected_site_rate_cards = v1.selectbox("Site Selection for Rate Cards :", site_list_rates, index=0)
        display_box_customers = v2.selectbox("Select Customers to View :", ["All", "Outliers Only"], index=0)
        st.text("")
        st.text("")

        p1_Storage, p2_Handling, = st.columns((1, 1))

        customer_rate_cards = customer_rate_cards[customer_rate_cards.Site == selected_site_rate_cards]

        with p1_Storage:
            df_pStorage = customer_rate_cards[customer_rate_cards.Description == "Storage"]
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
                if display_box_customers == "Outliers Only":
                    if float(y[i]) < q1_pStorage or float(y[i]) > q3_pStorage:
                        seperator = negative_separator if seperator == positive_separator else positive_separator
                        fig.add_annotation(
                            x=float(x[i]),
                            y=float(y[i]),
                            showarrow=False,
                            xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            text=f"{customer}",
                            font={"size": 13,
                                  "color": 'black'})

                        fig.update_layout(
                            title=dict(
                                text=f'{selected_site_rate_cards} : Customers ({df_pStorage_customers_count}) : Storage Rates'),
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
                        text=f"{customer}",
                        font={"size": 13,
                              "color": 'black'})

                    fig.update_layout(
                        title=dict(
                            text=f'{selected_site_rate_cards} : Customers ({df_pStorage_customers_count}) : Storage Rates'),
                        margin=dict(l=0, r=10, b=10, t=30),
                    )

            p1_Storage.plotly_chart(fig)

        # Handling Box plot ####################################################################
        with p2_Handling:
            df_pHandling = customer_rate_cards[customer_rate_cards.Description == "Pallet Handling"]
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
                            text=f"{customer}",
                            font={"size": 13,
                                  "color": 'black'}
                        )

                        fig.update_layout(
                            title=dict(
                                text=f'{selected_site_rate_cards} : Customers ({df_pHandling_customers_count}): Handling Rates'),
                            margin=dict(l=0, r=10, b=10, t=30)
                        )
                else:
                    seperator = negative_separator if seperator == positive_separator else positive_separator
                    fig.add_annotation(
                        x=float(x[i]),
                        y=float(y[i]),
                        showarrow=False,
                        xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                        hovertext=f'{customer}, Rate:  {val}',
                        text=f"{customer}",
                        font={"size": 13,
                              "color": 'black'})

                    fig.update_layout(
                        title=dict(
                            text=f'{selected_site_rate_cards} : Customers ({df_pHandling_customers_count}): Handling Rates'),
                        margin=dict(l=0, r=10, b=10, t=30)
                    )

            p2_Handling.plotly_chart(fig)

        st.divider()
        st.markdown("##### Carton Picking and Shrink Wrap Rates")
        st.divider()

        p3_Wrapping, p4_Cartons, = st.columns((1, 1))

        try:

            # Shrink Wrap Box plot ######################################################

            with p3_Wrapping:
                df_pWrapping = customer_rate_cards[customer_rate_cards.Description == "Shrink Wrap"]
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
                                text=f"{customer}",
                                font={"size": 13,
                                      "color": 'black'}
                            )

                            fig.update_layout(
                                title=dict(
                                    text=f'{selected_site_rate_cards} : Customers ({df_pWrapping_customers_count}) : Shrink Wrapping Rates'),
                                margin=dict(l=0, r=10, b=10, t=20))
                    else:
                        seperator = negative_separator if seperator == positive_separator else positive_separator
                        fig.add_annotation(
                            x=float(x[i]),
                            y=float(y[i]),
                            showarrow=False,
                            xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            text=f"{customer}",
                            font={"size": 13,
                                  "color": 'black'})

                        fig.update_layout(
                            title=dict(
                                text=f'{selected_site_rate_cards} : Customers ({df_pWrapping_customers_count}) : Shrink Wrapping Rates'),
                            margin=dict(l=0, r=10, b=10, t=20))

                p3_Wrapping.plotly_chart(fig)

            with p4_Cartons:
                ############################ Shrink Wrap Box plot ######################################################
                df_pCarton = customer_rate_cards[customer_rate_cards.Description == "Carton Picking"]
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
                                text=f"{customer}",
                                font={"size": 13,
                                      "color": 'black'}
                            )

                            fig.update_layout(
                                title=dict(
                                    text=f'{selected_site_rate_cards} : Customers ({df_pCarton_customers_count}) : Carton Picking Rates'),
                                margin=dict(l=0, r=10, b=10, t=20))
                    else:
                        seperator = negative_separator if seperator == positive_separator else positive_separator
                        fig.add_annotation(
                            x=float(x[i]),
                            y=float(y[i]),
                            showarrow=False,
                            xshift=float(y[i]) + seperator,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            text=f"{customer}",
                            font={"size": 13,
                                  "color": 'black'}
                        )

                        fig.update_layout(
                            title=dict(
                                text=f'{selected_site_rate_cards} : Customers ({df_pCarton_customers_count}) : Carton Picking Rates'),
                            margin=dict(l=0, r=10, b=10, t=20))

                p4_Cartons.plotly_chart(fig)

        except Exception as e:
            st.markdown(f"##############  Data pulled for this Customer or Site does not have all the Columns needed "
                        f"to display info ")
    with SiteTab:
        ############################### Side for Uploading excel File ###############################

        cost_centres = invoice_rates[invoice_rates.Cost_Center.str.contains("S&H", na=False)]
        cost_centres = cost_centres.Cost_Center.unique()
        all_workday_Customer_names = invoice_rates.WorkdayCustomer_Name.unique()

        select_Option1, select_Option2 = st.columns(2)

        with select_Option1:
            selected_cost_centre = st.selectbox("Cost Centre :", cost_centres, index=0)
            selected_workday_customers = invoice_rates[
                invoice_rates.Cost_Center == selected_cost_centre].WorkdayCustomer_Name.unique()
            select_CC_data = invoice_rates[invoice_rates.Cost_Center == selected_cost_centre]

        with select_Option2:
            if selected_cost_centre:
                select_customer = st.selectbox("Select Customer : ", selected_workday_customers, index=0)
            else:
                select_customer = st.selectbox("Select Customer : ", all_workday_Customer_names, index=0)

        st.divider()

        ###############################  Details Tab ###############################

        Workday_Sales_Item_Name = invoice_rates.Workday_Sales_Item_Name.unique()
        selected_customer = select_CC_data[select_CC_data.WorkdayCustomer_Name == select_customer]

        Revenue_Category = selected_customer.Revenue_Category.dropna().unique()

        selected_customer_pivot = pd.pivot_table(selected_customer,
                                                 values=["Quantity", "LineAmount"],
                                                 index=["Revenue_Category", "Workday_Sales_Item_Name"],
                                                 aggfunc="sum").reset_index()
        selected_customer_pivot[
            "Avg Rate"] = selected_customer_pivot.LineAmount / selected_customer_pivot.Quantity

        col1, col2 = st.columns(2)

        with col1:

            customer_avg_plts = selected_customer[selected_customer['Revenue_Category'] == 'Storage - Renewal'][
                "Quantity"].mean()
            site_avg_plts_occupied = select_CC_data[select_CC_data['Revenue_Category'] == 'Storage - Renewal'][
                "Quantity"].mean()

            selected_customer_pivot_table = selected_customer_pivot.style.format({
                "LineAmount": "${0:,.2f}",
                "Quantity": "{0:,.0f}",
                "Avg Rate": "${0:,.2f}",
            })

            selected_customer_pivot_table = selected_customer_pivot_table.map(highlight_negative_values)

            selected_customer_pivot_table = style_dataframe((selected_customer_pivot_table))
            selected_customer_pivot_table = selected_customer_pivot_table.hide(axis="index")

            st.write(selected_customer_pivot_table.to_html(), unsafe_allow_html=True)

            # Convert DataFrame to Excel
            output1567 = io.BytesIO()
            with pd.ExcelWriter(output1567) as writer:
                selected_customer_pivot.to_excel(writer)

            # Create a download button
            st.download_button(
                label="Download Excel  ⤵️",
                data=output1567,
                file_name='invoiced_data.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=1577
            )

            kpi1, kpi2, kpi3, = st.columns(3)

            with kpi1:
                # delta_value = (customer_avg_plts / site_avg_plts_occupied)/100
                # formartted_value = "{:.2%}".format(delta_value)
                st.metric(label=" Customer Avg Recurring Storage", value=f"{customer_avg_plts:,.0f} Plts",
                          # delta=f"~{formartted_value} of CC Holding",
                          border=True)

            with kpi2:
                pallet_billed_services = ['Blast Freezing', 'Storage - Initial', 'Storage - Renewal']

                filtered_data = \
                    selected_customer[selected_customer["Revenue_Category"].isin(pallet_billed_services)][
                        "LineAmount"].sum()
                ancillary_data = selected_customer["LineAmount"].sum() - filtered_data

                filtered_data_contribution = filtered_data / (filtered_data + ancillary_data)

                st.metric(label="Rent & Storage & Blast", value=f"${filtered_data:,.0f} ",
                          delta=f"{filtered_data_contribution:,.0%} : of Revenue", border=True)

            with kpi3:
                st.metric(label="Services", value=f"${ancillary_data:,.0f}",
                          delta=f"{1 - filtered_data_contribution:,.0%} : of Revenue", border=True)

        with col2:
            try:
                selected_graph = st.selectbox("Display :", ["Revenue", "Volumes"], index=0)
                graph_values = selected_graph

                if selected_graph == "Revenue":
                    graph_values = "LineAmount"
                else:
                    graph_values = "Quantity"

                selected_customer_pivot_pie_chart = selected_customer_pivot.drop(["Energy Surcharge",
                                                                                  "Other - Delayed Pallet Hire Revenue",
                                                                                  "Other - 3rd Party Recharge Revenue"]
                                                                                 , axis='index', errors='ignore')

                fig = px.pie(selected_customer_pivot_pie_chart, values=selected_customer_pivot_pie_chart[graph_values],
                             names=selected_customer_pivot_pie_chart["Revenue_Category"],
                             title=f'{selected_graph} View',
                             height=300, width=200)
                fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), )
                st.plotly_chart(fig, use_container_width=True)

            except Exception:
                st.markdown(
                    '#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
                output955 = io.BytesIO()
                with pd.ExcelWriter(output955) as writer:
                    selected_customer_pivot_pie_chart.to_excel(writer, sheet_name='export_data', index=False)

                # Create a download button
                st.download_button(
                    label="👆 Download ⤵️",
                    data=output955,
                    file_name='customer_revenue.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    key=969,
                )

        st.divider()

        ###############################  Pallet Positions vs Thru Put Pallets  ###############################
        try:

            select_CC_data_tab1 = select_CC_data[select_CC_data.WorkdayCustomer_Name == select_customer]
            storage_holding = select_CC_data_tab1[select_CC_data_tab1['Revenue_Category'] == 'Storage - Renewal']
            storage_holding.set_index("InvoiceDate", inplace=True)
            storage_holding = storage_holding.Quantity

            pallet_Thru_Put = select_CC_data_tab1[(select_CC_data_tab1.Revenue_Category == 'Handling - Initial') |
                                                  (select_CC_data_tab1.Revenue_Category == 'Handling Out')]

            pallet_Thru_Put_pivot = pd.pivot_table(pallet_Thru_Put,
                                                   values=["Quantity"],
                                                   index=["InvoiceDate"],
                                                   aggfunc="sum")

            fig2 = go.Figure()
            fig2.add_trace(
                go.Bar(x=pallet_Thru_Put_pivot.index, y=storage_holding, name="Pallet Holding"))

            fig2.add_trace(go.Scatter(x=pallet_Thru_Put_pivot.index, y=pallet_Thru_Put_pivot.Quantity, mode="lines",
                                      name="Plt Thru Put", yaxis='y2'))
            fig2.update_layout(
                title="Pallet Position vs Thru Put Pallets",
                xaxis=dict(title=select_customer),
                yaxis=dict(title="Pallets Stored", showgrid=False),
                yaxis2=dict(title="Pallet Thru Put", overlaying='y', side="right"),
                template="gridon",
                legend=dict(x=1, y=1)

            )

            st.plotly_chart(fig2, use_container_width=True)
        except Exception:
            st.markdown(
                '#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
            output1006 = io.BytesIO()
            with pd.ExcelWriter(output1006) as writer:
                select_CC_data_tab1.to_excel(writer, sheet_name='export_data', index=False)

            # Create a download button
            st.download_button(
                label="👆 Download ⤵️",
                data=output1006,
                file_name='data_download.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=1017,
            )

        st.divider()
        sdiv, sdiv2, sdiv3 = st.columns((1, 4, 1))
        sdiv.write("")

        ###############################  Box Plots Over and under Indexed Customers  ###############################

        sdiv2.subheader('Outliers View:- Under vs Over Indexed Customer Rates')
        sdiv3.markdown("###### Based on 2024 Invoiced data.... (Filter out Outliers)")
        st.divider()

        p1_Storage, p2_Handling, p3_Wrapping = st.columns((1, 1, 1))

        try:
            with p1_Storage:
                df_pStorage = \
                    select_CC_data[(select_CC_data['Revenue_Category'] == 'Storage - Initial') |
                                   (select_CC_data['Revenue_Category'] == 'Storage - Renewal')
                                   ].groupby(
                        "WorkdayCustomer_Name")[
                        "UnitPrice"].mean()

                customer_names = df_pStorage.index
                q3_pStorage = np.percentile(df_pStorage.values, 75)
                q1_pStorage = np.percentile(df_pStorage.values, 25)

                # #### Storage Box plot ####################################################################
                # # Add points with Labels
                y = df_pStorage.values
                x = np.random.normal(1, 0.250, size=(len(y)))

                # Create Box Plot Figure:
                box_x_value = np.random.normal(1, 0.250, size=len(y))

                # Create a Boxplot
                fig = go.Figure()

                fig.add_trace(go.Box(
                    y=y,
                    boxpoints='all',
                    jitter=0.3,
                    pointpos=0.5,
                    name="Storage Rate",
                ))

                # Add annotations for each data point. to recheck
                for i, (val, customer) in enumerate(zip(y, customer_names)):
                    fig.add_annotation(
                        y=val,
                        showarrow=False,
                        xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
                        hovertext=f'{customer}, Rate:  {val}',
                        text="*",

                    )
                    if customer == select_customer:
                        customer = select_customer
                        fig.update_annotations(
                            y=val,
                            text=f'{customer}, Rate:  {val}',
                            showarrow=False,
                            xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            # name=f'Customer: {customer}, Rate:  {val}',
                        )

                fig.update_layout(
                    title=dict(text='Cost Centre Storage Rates'),
                    margin=dict(l=0, r=10, b=10, t=30),
                )

                p1_Storage.plotly_chart(fig)

        except Exception:
            st.markdown(
                '#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
            output1093 = io.BytesIO()
            with pd.ExcelWriter(output1093) as writer:
                df_pStorage.to_excel(writer, sheet_name='export_data', index=False)

            # Create a download button
            st.download_button(
                label="👆 Download ⤵️",
                data=output1093,
                file_name='data_download.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=1088,
            )

        try:

            with p2_Handling:

                #### Handling Box plot ####################################################################
                df_pHandling = select_CC_data[(select_CC_data['Revenue_Category'] == 'Handling - Initial') &
                                              (select_CC_data['Revenue_Category'] == 'Handling - Initial')].groupby(
                    "WorkdayCustomer_Name")["UnitPrice"].mean()

                handling_customer_names = df_pHandling.index
                q3_pHandling = np.percentile(df_pHandling.values, 75)
                q1_pHandling = np.percentile(df_pHandling.values, 25)

                # Add points with Labels
                y = df_pHandling.values
                x = np.random.normal(1, 0.250, size=(len(y)))

                # Create Box Plot Figure:
                box_x_value = np.random.normal(1, 0.250, size=len(y))

                # Create a Boxplot
                fig = go.Figure()

                fig.add_trace(go.Box(
                    y=y,
                    boxpoints='all',
                    jitter=0.3,
                    pointpos=0.5,
                    name="Handling Rates"
                ))

                # Add annotations for each data point
                for i, (val, customer) in enumerate(zip(y, handling_customer_names)):
                    fig.add_annotation(
                        y=val,
                        text="*",
                        showarrow=False,
                        xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
                        hovertext=f'{customer}, Rate:  {val}',
                    )
                    if customer == select_customer:
                        customer = select_customer
                        fig.update_annotations(
                            y=val,
                            text=f'{customer}, Rate:  {val}',
                            showarrow=False,
                            xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            # name=f'Customer: {customer}, Rate:  {val}',
                        )

                fig.update_layout(
                    title=dict(text='Spread of Customer Handling Rates'),
                    margin=dict(l=0, r=10, b=10, t=30)
                )

                p2_Handling.plotly_chart(fig)
        except Exception:
            st.markdown(
                '#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
            output1165 = io.BytesIO()
            with pd.ExcelWriter(output1165) as writer:
                df_pHandling.to_excel(writer, sheet_name='export_data', index=False)

            # Create a download button
            st.download_button(
                label="👆 Download ⤵️",
                data=output1165,
                file_name='data_download.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=1160,
            )

        try:

            with p3_Wrapping:
                ############################ Shrink Wrap Box plot ######################################################
                df_pWrapping = select_CC_data[select_CC_data['Revenue_Category'] == 'Accessorial - Shrink Wrap'] \
                    .groupby("WorkdayCustomer_Name")["UnitPrice"].mean()

                wrapping_customer_names = df_pWrapping.index
                q3_pWrapping = np.percentile(df_pWrapping.values, 75)
                q1_pWrapping = np.percentile(df_pWrapping.values, 25)

                # Add points with Labels
                y = df_pWrapping.values
                x = np.random.normal(1, 0.250, size=(len(y)))

                # Create Box Plot Figure:
                box_x_value = np.random.normal(1, 0.250, size=len(y))

                # Create a Boxplot
                fig = go.Figure()
                fig.add_trace(go.Box(
                    y=y,
                    boxpoints='all',
                    jitter=0.3,
                    pointpos=0.5,
                    name="Shrink Wrap Rates"
                ))

                # Add annotations for each data point
                for i, (val, customer) in enumerate(zip(y, wrapping_customer_names)):
                    fig.add_annotation(
                        y=val,
                        showarrow=False,
                        xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
                        hovertext=f'{customer}, Rate:  {val}',
                        text="*",
                    )
                    if customer == select_customer:
                        customer = select_customer
                        fig.update_annotations(
                            y=val,
                            text=f'{customer}, Rate:  {val}',
                            showarrow=False,
                            xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
                            hovertext=f'{customer}, Rate:  {val}',
                            # name=f'Customer: {customer}, Rate:  {val}',
                        )

                fig.update_layout(
                    title=dict(text='Spread of Customer Pallet Wrapping Rates'),
                    margin=dict(l=0, r=10, b=10, t=20))

                p3_Wrapping.plotly_chart(fig)
        except Exception as e:
            st.markdown(
                '###### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
            output1233 = io.BytesIO()
            with pd.ExcelWriter(output1233) as writer:
                df_pWrapping.to_excel(writer, sheet_name='export_data', index=False)

            # Create a download button
            st.download_button(
                label="👆 Download ⤵️",
                data=output1233,
                file_name='data_download.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=1228,
            )

        st.divider()
        try:
            selected_cost_Centre = select_CC_data[
                ["WorkdayCustomer_Name", "Revenue_Category", "LineAmount", "Quantity"]].fillna(0)

            selected_cost_Centre = selected_cost_Centre.groupby(["WorkdayCustomer_Name", "Revenue_Category"]
                                                                ).agg(
                {"LineAmount": 'sum', "Quantity": 'sum', }).reset_index()

            selected_LineAmount = selected_cost_Centre.pivot(index="WorkdayCustomer_Name",
                                                             columns="Revenue_Category",
                                                             values="LineAmount").reset_index()

            selected_LineAmount = selected_LineAmount.fillna(0)

            selected_LineAmount["Total"] = selected_LineAmount.iloc[:, 1:].sum(axis=1)

            selected_LineAmount = selected_LineAmount.nlargest(10, ["Total"])

            selected_LineAmount.reindex()

            columns = selected_LineAmount.columns

            #### Plotting HEAT MAP - tow Show Revenue and Profitabiliy #########################################

            st.markdown('##### Avg Recurring Pallet Storage Billed')
            st.divider()

            selected_customer_vol_billed = pd.pivot_table(select_CC_data,
                                                          values=["Quantity"],
                                                          index=["Revenue_Category", "WorkdayCustomer_Name"],
                                                          aggfunc="mean")

            selected_customer_vol_billed = selected_customer_vol_billed.query(
                "WorkdayCustomer_Name != '.All Other [.All Other]'")
            treemap_data = selected_customer_vol_billed.loc["Storage - Renewal"]
            customers = treemap_data.index

            fig = px.treemap(treemap_data,
                             path=[customers],
                             values=treemap_data["Quantity"],
                             color=treemap_data["Quantity"],
                             custom_data=[customers],
                             color_continuous_scale='RdYlGn',

                             # title="Testing",
                             )

            st.plotly_chart(fig, filename='Chart.html')
            st.divider()

            selected_LineAmount.rename(
                columns={0: "Other"},
                inplace=True)

            sortedListed = selected_LineAmount.columns.sort_values(ascending=False)
            download_data = selected_LineAmount[list(sortedListed)]
            st.data_editor(download_data)

            output1305 = io.BytesIO()
            with pd.ExcelWriter(output1305) as writer:
                treemap_data.to_excel(writer, sheet_name='export_data', index=False)

            # Create a download button
            st.download_button(
                label="Download file️ ⤵️",
                data=output1305,
                file_name='customer_revenue.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=2007,
            )
        except Exception as e:
            st.markdown(
                '#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
            output1319 = io.BytesIO()
            with pd.ExcelWriter(output1319) as writer:
                select_CC_data.to_excel(writer, sheet_name='export_data', index=False)

            # Create a download button
            st.download_button(
                label="👆 Download ⤵️",
                data=output1319,
                file_name='data_download.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=1316,
            )

        st.divider()
    with AboutTab:
        st.markdown(f'##### Arthur Rusike: *Report any Code breaks or when you see Red Error Message*'
                    )
