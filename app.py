import io
import random
from io import BytesIO
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from helper_functions import style_dataframe, load_profitbility_Summary_model,load_rates_standardisation, load_specific_xls_sheet

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

    except Exception as e:

        st.subheader('To use with this WebApp 📊  - Upload below : \n '
                     '1. Profitability Summary File \n'
                     '2. Customer Rates Summary File ')

# Main Dashboard ##############################

if uploaded_file and customer_rates_file:
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
    customer_rate_cards = load_specific_xls_sheet(customer_rates_file, "2025_Rate_Cards", 0, "A:I")

    # Read individual Excel sheet separately = Customer Volumes current year and prior year. Use for comparison.
    dominic_report = load_specific_xls_sheet(customer_rates_file, "dominic_report", 0, "A:AM")

    # Customer Names for Selection in Select Box
    all_workday_customer_names = profitability_summary_file.Customer.unique()

    # Cost Centre Names for Selection in Select Box
    cost_centres = profitability_summary_file["Cost Center"].unique()

    # Site Names for Selection in Select Box
    site_list = list(profitability_summary_file.Site.unique())

    space_holder_1, title_holder, cost_centres_selection, space_holder3 = st.columns(4)

    title_holder.subheader("Project Renaissance", divider="blue")
    selected_site = cost_centres_selection.multiselect("Site :", site_list, site_list[1])

    DetailsTab, SiteTab, ProfitabilityTab, AboutTab = st.tabs(["📊 Quick View", "🥇 Site Detailed View",
                                                               "📈 Estimate Profitability",
                                                               "ℹ️ About"])
    #  Details Tab ###############################

    with DetailsTab:

        select_Site_data = profitability_summary_file[profitability_summary_file.Site.isin(selected_site)]

        select_Site_data_2023_profitability = profitability_2023[
            profitability_2023.Site.isin(selected_site)]
        select_Site_data_2023_pallets = customer_pallets[customer_pallets.Site.isin(selected_site)]

        kpi1, kpi2, kpi3, kpi4, = st.columns(4)

        with kpi1:
            revenue_2024 = select_Site_data[" Revenue"].sum() / 1000
            revenue_2023 = select_Site_data_2023_profitability[" Revenue"].sum() / 1000
            delta = f"{((revenue_2024 / revenue_2023) - 1):,.1%}"

            kpi1.metric(label=f"FY24: Revenue - 000s", value=f"{revenue_2024:,.0f}", delta=f"{delta} vs LY ")

        with kpi2:
            ebitda_2024 = select_Site_data["EBITDA $"].sum() / 1000
            ebitda_2023 = select_Site_data_2023_profitability["EBITDA $"].sum() / 1000
            delta = f"{((ebitda_2024 / ebitda_2023) - 1):,.1%}"

            ebitda_margin = f"{(ebitda_2024 / revenue_2024):,.1%}"

            kpi2.metric(label=f"FY24: EBITDA - 000s", value=f"{ebitda_2024:,.0f}",
                        delta=f"{delta} | FY24 → {ebitda_margin}   ")

        with kpi3:

            dominic_report_2024 = dominic_report[(dominic_report.Year == 2024) & (dominic_report.Month == "December")]
            dominic_report_2024["Site"] = dominic_report_2024["Cost Centers"].apply(extract_site)
            dominic_report_2024 = dominic_report_2024[dominic_report_2024.Site.isin(selected_site)]

            dominic_report_2023 = dominic_report[(dominic_report.Year == 2023) & (dominic_report.Month == "December")]
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

        with kpi4:
            volume_guarantee_2024 = select_Site_data_2023_pallets["VG Pallets - Dec-2024"].sum()
            volume_guarantee_2023 = select_Site_data_2023_pallets["VG Pallets - Dec-2023"].sum()

            delta = f"{(volume_guarantee_2024 / capacity_2024):,.1%} : of Capacity is VG "

            kpi4.metric(label=f"Volume Guarantee", value=f"{volume_guarantee_2024:,.0f}", delta=f"{delta} ")

        st.divider()


        def highlight_negative_values(value):
            if type(value) != str:
                if value < 0:
                    return f"color: red"


        fin1, fin2 = st.columns(2)

        key_metrics_data = profitability_summary_file[profitability_summary_file[" Revenue"] > 1]
        key_metrics_data = key_metrics_data[key_metrics_data.Site.isin(selected_site)]
        key_metrics_data["Name"] = key_metrics_data["Cost Center"].apply(extract_cost_name)

        key_metrics_data_pivot = pd.pivot_table(key_metrics_data, index=["Cost Center", "Name"],
                                                values=[" Revenue", 'EBITDAR $', 'Rent Expense,\n$',
                                                        'EBITDA $'
                                                        ],
                                                aggfunc='sum')

        key_metrics_data_pivot = key_metrics_data_pivot.reset_index("Name")

        key_metrics_data_pivot = key_metrics_data_pivot.style.hide(axis="index")

        key_metrics_data_pivot = key_metrics_data_pivot.format({
            " Revenue": '${0:,.0f}',
            'EBITDAR $': "${0:,.0f}",
            'Rent Expense,\n$': "${0:,.0f}",
            'EBITDA $': "${0:,.0f}",
        })

        key_metrics_data_pivot = key_metrics_data_pivot.map(highlight_negative_values)

        key_metrics_data_pivot = style_dataframe(key_metrics_data_pivot)

        with fin1:
            fin1.markdown("##### Financial Summary")
            fin1.write(key_metrics_data_pivot.to_html(), unsafe_allow_html=True )

        key_metrics_data = key_metrics_data.groupby("CC").sum()

        DL_Handling = key_metrics_data["Direct Labor Expense,\n$"]
        Ttl_Labour = key_metrics_data["Total Labor Expense, $"]
        Service_Revenue = key_metrics_data["Service Rev (Handling+ Case Pick+ Other Rev)"]

        key_metrics_data["Rev Per Plt"] = key_metrics_data[" Revenue"] / key_metrics_data["Pallet"]
        key_metrics_data["EBITDA Per Plt"] = key_metrics_data["EBITDA $"] / key_metrics_data["Pallet"]
        key_metrics_data["Stock Turn Times"] = key_metrics_data["Pallet"] / key_metrics_data["TTP p.w."]
        key_metrics_data["DL %"] = DL_Handling / Service_Revenue
        key_metrics_data["LTR %"] = Ttl_Labour / key_metrics_data[" Revenue"]

        kpi_key_metrics_data_pivot = key_metrics_data[
            ["Rev Per Plt", "EBITDA Per Plt", "Stock Turn Times", "DL %", "LTR %"]]

        kpi_key_metrics_data_pivot = kpi_key_metrics_data_pivot.style.hide(axis="index")

        kpi_key_metrics_data_pivot = kpi_key_metrics_data_pivot.format({
            'Rev Per Plt': "${0:,.2f}",
            'EBITDA Per Plt': "${0:,.2f}",
            'Stock Turn Times': "{0:,.2f}",
            'DL %': '{0:,.1%}',
            'LTR %': '{0:,.1%}',
        })

        kpi_key_metrics_data_pivot = kpi_key_metrics_data_pivot.map(highlight_negative_values)

        kpi_key_metrics_data_pivot = style_dataframe(kpi_key_metrics_data_pivot)

        with fin2:
            fin2.markdown("##### KPI Summary")
            fin2.write(kpi_key_metrics_data_pivot.to_html(), unsafe_allow_html=True  )

        st.divider()

        # TOP 10 Customers ##################################################

        st.subheader("Top 10 Customers")

        s1, s2, s3, s4 = st.columns(4)

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
        treemap_data = select_CC_data.filter(
            ['Name', " Revenue", 'EBITDA $', 'EBITDA Margin\n%', 'EBITDA per Pallet', 'Revenue per Pallet',
             'Pallet', 'TTP p.w.', 'Rank'])

        display_data = treemap_data
        size = len(display_data)
        display_data = display_data.query(f"Rank > {size - 10} ")
        display_data_pie = display_data
        rank_display_data = display_data
        display_data = display_data.style.hide(axis="index")
        display_data = display_data.hide(["Rank"], axis="columns")

        display_data = style_dataframe(display_data)

        display_data = display_data.format({
            " Revenue": '${0:,.0f}',
            'EBITDA $': "${0:,.0f}",
            'EBITDA Margin\n%': "{0:,.2%}",
            'Revenue per Pallet': "${0:,.2f}",
            'EBITDA per Pallet': "${0:,.2f}",
            'Pallet': "{0:,.0f}",
            'TTP p.w.': '{0:,.0f}'
        })

        display_data = display_data.map(highlight_negative_values)
        display_data.background_gradient(subset=['EBITDA Margin\n%'], cmap="RdYlGn")

        # Create a download button
        with s3:
            output1 = io.BytesIO()
            with pd.ExcelWriter(output1) as writer:
                display_data.to_excel(writer)

            st.text("")
            s3.download_button(
                label="👆 Download ⤵️",
                data= output1,
                file_name='customer_revenue.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=295,
            )

        with s4:
            selected_graph = s4.selectbox("Display :", [" Revenue", 'EBITDA $', 'EBITDA Margin\n%', 'EBITDA per Pallet',
                                                        'Revenue per Pallet',
                                                        'Pallet', 'TTP p.w.'], index=0)

        d1, d4 = st.columns((2, 1))

        with d1:
            d1.write(display_data.to_html(),unsafe_allow_html=True, use_container_width=True )

        with d4:

            fig = px.pie(display_data_pie, values=display_data_pie[selected_graph],
                         names=display_data_pie.Name,
                         title=f'{selected_graph} View',
                         height=300, width=200)
            fig.update_layout(margin=dict(l=20, r=20, t=30, b=0), )

            d4.plotly_chart(fig, use_container_width=True)

        with st.expander("Top 10 Customers - Activity Rank", expanded=True):

            rank1, rank2 = st.columns((2,1))
            with rank1:

                rank_display_data["Stock Turn Times"] = rank_display_data["Pallet"] / rank_display_data["TTP p.w."]
                rank_display_data["Rev Rank"] = rank_display_data[" Revenue"].rank(method='max', ascending=False)
                rank_display_data["EBITDA"] = rank_display_data["EBITDA $"].rank(method='max', ascending=False)
                rank_display_data["Margin %"] = rank_display_data["EBITDA Margin\n%"]
                rank_display_data["Rev | Pallet"] = rank_display_data["Revenue per Pallet"].rank(method='max',
                                                                                                 ascending=False)
                rank_display_data["EBITDA | Plt"] = rank_display_data["EBITDA per Pallet"].rank(method='max',
                                                                                                ascending=False)
                rank_display_data["Turns"] = rank_display_data["Stock Turn Times"].rank(method='max', ascending=False)
                rank_display_data["Pallets"] = rank_display_data["Pallet"].rank(method='max', ascending=False)
                rank_display_data["TTP"] = rank_display_data["TTP p.w."].rank(method='max', ascending=False)
                rank_display_data["Score"] = ((+ rank_display_data["EBITDA"] \
                                               + rank_display_data["Margin %"] + rank_display_data["Rev | Pallet"] \
                                               + rank_display_data["EBITDA | Plt"] + rank_display_data["Turns"] \
                                               + rank_display_data["Pallets"] + rank_display_data["TTP"]) / 7) / \
                                             rank_display_data["Rev Rank"]

                def scoring_customer(score1):
                    if score1 < 1000:
                        return 5
                    return rank_display_data["Score"]



                rank_display_data["Score"] = [ 5 if y < 1000 else x for x,y in zip(rank_display_data["Score"],rank_display_data["EBITDA $"])]


                def assign_score_card(score):
                    if score <= 0.50:
                        return "Excellent"
                    elif score <= 0.70:
                        return "Good"
                    elif score <= 1:
                        return "Satisfactory"
                    elif score <= 1.25:
                        return "Red Flag"
                    elif score <= 1.50:
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


                rank_display_data["Comment"] = rank_display_data["Score Card"].apply(add_comment)

                columns_to_hide = [" Revenue", 'EBITDA $', 'EBITDA Margin\n%', 'EBITDA per Pallet', 'Revenue per Pallet',
                                   'Pallet', 'TTP p.w.', 'Stock Turn Times']

                rank_display_data = rank_display_data.style.hide(axis="index")
                rank_display_data = rank_display_data.hide([" Revenue", 'EBITDA $', 'EBITDA Margin\n%',
                                                            'EBITDA per Pallet', 'Revenue per Pallet', 'Pallet', 'TTP p.w.',
                                                            'Stock Turn Times', 'Rank'], axis="columns")

                rank_display_data = rank_display_data.format({
                    "Rev Rank": '{0:,.0f}',
                    "EBITDA": '{0:,.0f}',
                    "Margin %": '{0:,.2%}',
                    "Rev | Pallet": '{0:,.0f}',
                    "EBITDA | Plt": '{0:,.0f}',
                    "Turns": '{0:,.0f}',
                    "Pallets": '{0:,.0f}',
                    "TTP": '{0:,.0f}',
                    "Score": '{0:,.3f}',
                })

                rank_display_data.background_gradient(subset=['Score'], cmap="RdYlGn_r")
                # def assign_score_card():
                #    score =   rank_display_data["Rev Rank"] +

                rank_display_data = style_dataframe(rank_display_data)
                rank_display_data = rank_display_data.map(highlight_negative_values)

                rank1.write(rank_display_data.to_html(), unsafe_allow_html=True, use_container_width=True )

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
                    key=367,
                )

            with rank2:
                rank2.markdown(f"""
                    __Score Card Definitions__ \n
                    Score - is Performance in rank of Customer's KPI relative to it's Revenue contribution rank. \n
                    Score Card : Scoring bands with Respective Comments\n
                    
                    |Band                      |    Definition                       |
                    |:-------------------------|:------------------------------------|
                    |+1.5 : High Risk          | Needs immediate review and action   |
                    |1.50 : Needs Improvement  | Identify areas of Improvement       |
                    |1.25 : Red Flag           | Future problem review now           |
                    |1.00 : Satisfactory       | Satisfactory to have Customer       |
                    |0.70 : Good               | Good to have Customer               |
                    |0.50 : Excellent          | Excellent to have Customer          |
                  
                """)

        st.divider()

        st.subheader(f'Revenue vs Profitability Contribution - Analysis : {selected_cost_centre}', divider='rainbow')

        graph1, graph2 = st.columns(2)

        customers = treemap_data['Name']
        # sales = [" Revenue", 'EBITDA $', 'EBITDA Margin\n%', 'Revenue per Pallet', 'EBITDA per Pallet', 'Pallet',
        #          'TTP p.w.']
        color_columns = [" Revenue", 'EBITDA $', 'EBITDA Margin\n%', 'Revenue per Pallet', 'EBITDA per Pallet',
                         'Pallet',
                         'TTP p.w.']
        remark = select_CC_data['EBITDA Margin\n%']  # Replace with EBITDA per Pallet Through Put
        margin = select_CC_data['EBITDA Margin\n%']  # Replace with EBITDA per Pallet Through Put

        with graph1:
            graph1_size, graph1_color, space_holder_3, space_holder_4 = st.columns(4)

            box_size_view = graph1_size.selectbox("Select Option for Box Size:", color_columns, key=111, index=0)
            box_color_view = graph1_color.selectbox("Select Option for Box Color:", color_columns, key=112, index=3)

            # Plotting HEAT MAP - tow Show Revenue and Profitability ############################

            fig = px.treemap(select_CC_data,
                             path=[customers],
                             values=select_CC_data[box_size_view],
                             color=select_CC_data[box_color_view],
                             color_continuous_scale='RdYlGn',
                             )

            st.plotly_chart(fig, filename='Chart.html')

        with graph2:
            treemap_data = treemap_data.filter(['Name', " Revenue", 'EBITDA $', 'EBITDA Margin\n%',
                                                'Revenue per Pallet', 'EBITDA per Pallet', 'Pallet',
                                                'TTP p.w.'])
            with st.expander("Filter Customer to Display"):
                st.data_editor(treemap_data, num_rows='dynamic')

            treemap_data = treemap_data[(treemap_data[box_size_view] > 1) & (treemap_data[box_color_view] > 1)]
            x_value = str(box_size_view)
            y_value = str(box_color_view)

            fig = px.scatter(treemap_data, x=y_value,
                             y=x_value,
                             size=y_value, color="Name",
                             hover_name="Name", log_x=False, size_max=60)
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
                marker={"size": 4},

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
                            title=dict(text=f'{selected_site_rate_cards} : Customers ({df_pStorage_customers_count}) : Storage Rates'),
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
                        title=dict(text=f'{selected_site_rate_cards} : Customers ({df_pStorage_customers_count}) : Storage Rates'),
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
                marker={"size": 4},
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
                            title=dict(text=f'{selected_site_rate_cards} : Customers ({df_pHandling_customers_count}): Handling Rates'),
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
                        title=dict(text=f'{selected_site_rate_cards} : Customers ({df_pHandling_customers_count}): Handling Rates'),
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
                    marker={"size": 4},
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
                                title=dict(text=f'{selected_site_rate_cards} : Customers ({df_pWrapping_customers_count}) : Shrink Wrapping Rates'),
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
                            title=dict(text=f'{selected_site_rate_cards} : Customers ({df_pWrapping_customers_count}) : Shrink Wrapping Rates'),
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
                    marker={"size": 4},
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
                                title=dict(text=f'{selected_site_rate_cards} : Customers ({df_pCarton_customers_count}) : Carton Picking Rates'),
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
                            title=dict(text=f'{selected_site_rate_cards} : Customers ({df_pCarton_customers_count}) : Carton Picking Rates'),
                            margin=dict(l=0, r=10, b=10, t=20))

                p4_Cartons.plotly_chart(fig)

        except Exception as e:
            st.markdown(f"##############  Data pulled for this Customer or Site does not have all the Columns needed "
                        f"to display info ")
    with SiteTab:
        ############################### Side for Uploading excel File ###############################

        with st.spinner('Loading and Updating Report...🥱'):
            try:
                with st.sidebar:
                    st.title("Site Detailed View")
                    uploaded_file = st.file_uploader('Upload Customer Rates File', type=['xlsx', 'xlsm', 'xlsb'])
                    # Main Dashboard
                    invoice_rates = load_invoices_model(uploaded_file)
                    st.divider()

                    st.markdown("#### Dates Covered by Invoice File")
                    start_date = invoice_rates.formatted_date.dropna()
                    st.markdown(f'###### Report Start Date  : {start_date.min()}')
                    st.markdown(f'###### Report End Date  : {start_date.max()}')
                    st.divider()

            except Exception as e:
                st.divider()
                st.subheader('Upload Customer Rates File to use with this App 📊 ')

        ###############################  Main Dashboard ##############################

        if uploaded_file:
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
                                                     aggfunc="sum")
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

                st.dataframe(selected_customer_pivot_table)

                # Convert DataFrame to Excel
                output891 = io.BytesIO()
                with pd.ExcelWriter(output891) as writer:
                    selected_customer_pivot.to_excel(writer)

                # Create a download button
                st.download_button(
                    label="Download Excel  ⤵️",
                    data=output891,
                    file_name='invoiced_data.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    key=210
                )

                kpi1, kpi2, kpi3, = st.columns(3)

                with kpi1:
                    # delta_value = (customer_avg_plts / site_avg_plts_occupied)/100
                    # formartted_value = "{:.2%}".format(delta_value)
                    st.metric(label=" Customer Avg Recurring Storage", value=f"{customer_avg_plts:,.0f} Plts",
                              # delta=f"~{formartted_value} of CC Holding",
                              border=True)

                with kpi2:
                    pallet_billed_services = ['Accessorial - Shrink Wrap', 'Blast Freezing', 'Handling - Initial',
                                              'Handling Out',
                                              'Storage - Initial', 'Storage - Renewal']

                    filtered_data = \
                        selected_customer[selected_customer["Revenue_Category"].isin(pallet_billed_services)][
                            "LineAmount"].sum()
                    ancillary_data = selected_customer["LineAmount"].sum() - filtered_data

                    st.metric(label="Billed Pallet Rev  ", value=f"${filtered_data:,.0f} ",
                              delta="excl Volume Guarantee ",
                              border=True)

                with kpi3:
                    st.metric(label="Billed Ancillary Rev", value=f"${ancillary_data:,.0f}",
                              border=True, )

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
                                                                                      "Other - 3rd Party Recharge Revenue" ]
                                                                                     ,axis='index', errors='ignore')

                    fig = px.pie(selected_customer_pivot_pie_chart, values=selected_customer_pivot_pie_chart[graph_values],
                                 names=selected_customer_pivot_pie_chart.index.get_level_values(0),
                                 title=f'{selected_graph} View',
                                 height=300, width=200)
                    fig.update_layout(margin=dict(l=20, r=20, t=30, b=0), )
                    st.plotly_chart(fig, use_container_width=True)

                except Exception:
                    st.markdown('#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
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
                st.markdown('#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
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
                st.markdown('#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
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
            except Exception :
                st.markdown('#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
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
                st.markdown('###### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
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
                    key=211,
                )
            except Exception as e:
                st.markdown('#### Data has Negative Values or is in a State that cannot be used with is Visualisation. Please review excel file')
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
else:
    st.subheader('To use with this WebApp 📊  - Upload below : \n '
                 '1. Profitability Summary File \n'
                 '2. Customer Rates Summary File ')
