# import streamlit as st
# import pandas as pd
#
#
# def sub_category_classification(dtframe):
#     Storage_Revenue = ['Storage - Initial', 'Storage - Renewal', 'Storage Guarantee']
#     Handling_Revenue = ['Handling - Initial', 'Handling Out']
#     Case_Pick_Revenue = ['Accessorial - Case Pick / Sorting']
#     Ancillary_Revenue = ['Accessorial - Documentation', 'Accessorial - Labeling / Stamping',
#                          'Accessorial - Labor and Overtime', 'Accessorial - Loading / Unloading / Lumping',
#                          'Accessorial - Palletizing', 'Accessorial - Shrink Wrap']
#     Blast_Freezing_Revenue = ['Blast Freezing', 'Room Freezing']
#     Other_Warehouse_Revenue = ['Other - Delayed Pallet Hire Revenue', 'Other - Warehouse Revenue',
#                                'Rental Electricity Income']
#
#     if dtframe["Revenue_Category"] in Storage_Revenue:
#         return "Storage Revenue"
#     elif dtframe["Revenue_Category"] in Handling_Revenue:
#         return "Handling Revenue"
#     elif dtframe["Revenue_Category"] in Case_Pick_Revenue:
#         return "Case Pick Revenue"
#     elif dtframe["Revenue_Category"] in Ancillary_Revenue:
#         return "Ancillary Revenue"
#     elif dtframe["Revenue_Category"] in Blast_Freezing_Revenue:
#         return "Blast Freezing Revenue"
#     elif dtframe["Revenue_Category"] in Other_Warehouse_Revenue:
#         return "Other Warehouse Revenue"
#     else:
#         return dtframe["Revenue_Category"]
#
#
# @st.cache_data
# def load_invoices_model(uploaded_file):
#     invoice_rates = pd.read_excel(uploaded_file, sheet_name="InvoiceRates")
#     invoice_rates["Calumo Description"] = invoice_rates.apply(sub_category_classification, axis=1)
#     invoice_rates['formatted_date'] = pd.to_datetime(invoice_rates['InvoiceDate'])
#     invoice_rates.formatted_date = invoice_rates.formatted_date.dt.strftime('%Y-%m-%d')
#     invoice_rates.sort_values("InvoiceNumber", inplace=True)
#     return invoice_rates
#
#
# @st.cache_data
# def profitability_model(uploaded_file):
#     AU_facility_financials = pd.read_excel(uploaded_file,
#                                            sheet_name="Facility Financials",
#                                            header=3,
#                                            usecols="A:EL")
#
#     return AU_facility_financials
#
#
#     # select_Site_data_2023_profitability = profitability_2023[
#     #     profitability_2023.Site.isin(selected_site)]
#     # select_Site_data_2023_pallets = customer_pallets[customer_pallets.Site.isin(selected_site)]
#
#     # with st.expander("Full Years Comparison - GAAP (USD) ", expanded=True):
#     #
#     #     section1, section2, section3 = st.columns((
#     #                                                            len(selected_site) * 0.20,
#     #                                                            len(selected_site) * 1,
#     #                                                            len(selected_site) * 1))
#     #
#     #     with section1:
#     #         selected_years = []
#     #
#     #         t1, t2, t3, t4, = st.columns((4, 1, 1, 1))
#     #
#     #         with t1:
#     #             st.text("")
#     #             st.text("")
#     #             t1.markdown("##### Years")
#     #             _2025 = t1.checkbox('2025', True)
#     #             _2024 = t1.checkbox('2024', True)
#     #             _2023 = t1.checkbox('2023', True)
#     #
#     #             if _2025:
#     #                 selected_years.append(2025)
#     #
#     #             if _2024:
#     #                 selected_years.append(2024)
#     #
#     #             if _2023:
#     #                 selected_years.append(2023)
#     #
#     #     budget_data_2025 = budget_data_2025[budget_data_2025.Year.isin(selected_years)]
#     #     budget_data_2025 = budget_data_2025[budget_data_2025.Site.isin(selected_site)]
#     #
#     #     budget_data_2025 = budget_data_2025.groupby(by=["Site", "Year"]).sum().reset_index()
#     #     budget_data_2025 = budget_data_2025.sort_values(by=["Site", "Year"], ascending=[False, False])
#     #
#     #     budget_data_2025_pivot = budget_data_2025
#     #     budget_data_2025.insert(0, "Description",
#     #                             budget_data_2025_pivot["Site"] + " : " + budget_data_2025_pivot["Year"].astype(
#     #                                 str))
#     #     budget_data_2025_pivot = budget_data_2025_pivot.set_index("Description")
#     #
#     #
#     #     def format_for_percentage(value):
#     #         return "{:.2%}".format(value)
#     #
#     #
#     #     def format_for_currency(value):
#     #         return "${0:,.0f}".format(value)
#     #
#     #
#     #     def format_for_float_currency(value):
#     #         return "${0:,.2f}".format(value)
#     #
#     #
#     #     def format_for_float(value):
#     #         return "{0:,.2f}".format(value)
#     #
#     #
#     #     def format_for_int(value):
#     #         return "{0:,.0f}".format(value)
#     #
#     #
#     #     budget_data_2025_pivot["Revenue"] = budget_data_2025_pivot["Revenue"].apply(format_for_currency)
#     #     budget_data_2025_pivot["Ebitda"] = budget_data_2025_pivot["Ebitda"].apply(format_for_currency)
#     #     budget_data_2025_pivot["Rev Per Pallet"] = budget_data_2025_pivot["Rev Per Pallet"].apply(
#     #         format_for_float_currency)
#     #     budget_data_2025_pivot["Ebitda Per Pallet"] = budget_data_2025_pivot["Ebitda Per Pallet"].apply(
#     #         format_for_float_currency)
#     #     budget_data_2025_pivot["Economic Utilization"] = budget_data_2025_pivot["Economic Utilization"].apply(
#     #         format_for_percentage)
#     #     budget_data_2025_pivot["Economic OHP"] = (budget_data_2025_pivot["Economic OHP"] / 12).apply(
#     #         format_for_int)
#     #     budget_data_2025_pivot["Labour to Tot. Rev"] = budget_data_2025_pivot["Labour to Tot. Rev"].apply(
#     #         format_for_percentage)
#     #     budget_data_2025_pivot["DL to Svcs Rev"] = budget_data_2025_pivot["DL to Svcs Rev"].apply(
#     #         format_for_percentage)
#     #     budget_data_2025_pivot["Direct Labor / Hour"] = budget_data_2025_pivot["Direct Labor / Hour"].apply(
#     #         format_for_float_currency)
#     #     budget_data_2025_pivot["Turn"] = budget_data_2025_pivot["Turn"].apply(format_for_float)
#     #     budget_data_2025_pivot["Rev - Storage & Blast"] = budget_data_2025_pivot["Rev - Storage & Blast"].apply(
#     #         format_for_percentage)
#     #     budget_data_2025_pivot["Rev - Services"] = budget_data_2025_pivot["Rev - Services"].apply(
#     #         format_for_percentage)
#     #
#     #     budget_data_2025_pivot = budget_data_2025_pivot.T
#     #
#     #     budget_data_2025_pivot = budget_data_2025_pivot[2:]
#     #
#     #     styles = [
#     #         {
#     #             'selector': ' tr:hover',
#     #             'props': [
#     #                 ('border', '1px solid #4CAF50'),
#     #                 ('background-color', 'wheat'),
#     #
#     #             ]
#     #         },
#     #         {
#     #             'selector': ' tr',
#     #             'props': [
#     #                 ('text-align', 'right'),
#     #                 ('font-size', '12px'),
#     #                 ('font-family', 'sans-serif, Arial'),
#     #
#     #             ]
#     #         }
#     #
#     #     ]
#     #
#     #     for column in budget_data_2025_pivot.columns:
#     #         styles.append({
#     #
#     #             'selector': f'th.col{budget_data_2025_pivot.columns.get_loc(column)}',
#     #             'props': [
#     #                 ('background-color', '#305496'),
#     #                 ('width', 'auto'),
#     #                 ('color', 'white'),
#     #                 ('font-family', 'sans-serif, Arial'),
#     #                 ('font-size', '12px'),
#     #                 ('text-align', 'center'),
#     #                 ('border', '2px solid white')
#     #             ],
#     #
#     #         }
#     #         )
#     #
#     #     budget_data_2025_pivot = budget_data_2025_pivot.style.set_table_styles(styles)
#     #
#     #     section2.markdown("##### Site Financial Summary - USD")
#     #
#     #     section2.write(budget_data_2025_pivot.to_html(), unsafe_allow_html=True)
#
#     # with section3:
#     #
#     #     bar1,_bar_dash, bar2 = st.columns((2,0.5,2))
#     #
#     #     budget_data_2025["Reve Per Pallet"] = budget_data_2025["Rev Per Pallet"] - budget_data_2025["Ebitda Per Pallet"]
#     #
#     #     def format_for_float_2_currency(value):
#     #         value = value/1000_000
#     #         return "${0:,.2f}M".format(value)
#     #
#     #
#     #     fig = px.bar(budget_data_2025, x='Year',
#     #                  y='Revenue',
#     #                  title=f"Site Revenue",
#     #                  text=budget_data_2025["Revenue"].map(format_for_float_2_currency)
#     #                  )
#     #     fig.update_traces(width=0.7,
#     #                       marker=dict(cornerradius=30),
#     #                       )
#     #     bar1.plotly_chart(fig,key=303 , use_container_width=True)
#     #
#     #     fig = px.bar(budget_data_2025, x='Year',
#     #                  y='Ebitda',
#     #                  title="Site Ebitda",
#     #                  text=budget_data_2025["Ebitda"].map(format_for_float_2_currency)
#     #
#     #                  )
#     #     fig.update_traces(width=0.5,
#     #                       marker=dict(cornerradius=30),
#     #                       )
#     #     bar2.plotly_chart(fig,key=308,  use_container_width=True )
#
#
#     # with section3:
#     #     st.text("")
#     #     st.text("")
#     #     st.text("")
#     #
#     #     kpi1, kpi2, = st.columns(2)
#     #
#     #     # with kpi1:
#     #     revenue_2024 = select_Site_data[" Revenue"].sum() / 1000
#     #     revenue_2023 = select_Site_data_2023_profitability[" Revenue"].sum() / 1000
#     #     delta = f"{((revenue_2024 / revenue_2023) - 1):,.1%}"
#     #
#     #     kpi1.metric(label=f"FY24: Revenue - 000s", value=f"{revenue_2024:,.0f}", delta=f"{delta} vs LY ")
#     #
#     #     # with kpi2:
#     #     ebitda_2024 = select_Site_data["EBITDA $"].sum() / 1000
#     #     ebitda_2023 = select_Site_data_2023_profitability["EBITDA $"].sum() / 1000
#     #     delta = f"{((ebitda_2024 / ebitda_2023) - 1):,.1%}"
#     #
#     #     ebitda_margin = f"{(ebitda_2024 / revenue_2024):,.1%}"
#     #
#     #     kpi2.metric(label=f"FY24: EBITDA - 000s", value=f"{ebitda_2024:,.0f}",
#     #                 delta=f"{delta} | FY24 → {ebitda_margin}   ")
#     #     st.text("")
#     #     st.text("")
#     #
#     #     kpi3, kpi4, = st.columns(2)
#     #
#     #     dominic_report_2024 = dominic_report[
#     #         (dominic_report.Year == 2024) & (dominic_report.Month == "December")]
#     #     dominic_report_2024["Site"] = dominic_report_2024["Cost Centers"].apply(extract_site)
#     #     dominic_report_2024 = dominic_report_2024[dominic_report_2024.Site.isin(selected_site)]
#     #
#     #     dominic_report_2023 = dominic_report[
#     #         (dominic_report.Year == 2023) & (dominic_report.Month == "December")]
#     #     dominic_report_2023["Site"] = dominic_report_2023["Cost Centers"].apply(extract_site)
#     #     dominic_report_2023 = dominic_report_2023[dominic_report_2023.Site.isin(selected_site)]
#     #
#     #     economic_pallets_2024 = dominic_report_2024["Economic Pallet Total"].sum()
#     #     capacity_2024 = dominic_report_2024.Denominator.sum()
#     #
#     #     economic_pallets_2023 = dominic_report_2023["Economic Pallet Total"].sum()
#     #     capacity_2023 = dominic_report_2023.Denominator.sum()
#     #
#     #     dec_2024_occupancy = economic_pallets_2024 / capacity_2024
#     #     dec_2023_occupancy = economic_pallets_2023 / capacity_2023
#     #
#     #     delta = f"{((dec_2024_occupancy / dec_2023_occupancy) - 1):,.1%}"
#     #
#     #     kpi3.metric(label=f"Economic Occupancy", value=f"{dec_2024_occupancy:,.1%}", delta=f"{delta} ")
#     #
#     #     # with kpi4:
#     #     volume_guarantee_2024 = select_Site_data_2023_pallets["VG Pallets - Dec-2024"].sum()
#     #     volume_guarantee_2023 = select_Site_data_2023_pallets["VG Pallets - Dec-2023"].sum()
#     #
#     #     delta = f"{(volume_guarantee_2024 / capacity_2024):,.1%} : of Capacity is VG "
#     #
#     #     kpi4.metric(label=f"Volume Guarantee", value=f"{volume_guarantee_2024:,.0f}", delta=f"{delta} ")
#     # st.text("")
#     # st.text("")
#
#     # with st.expander("2024 Financial and KPI Summaries - Adjusted for Market Rent", expanded=False):
#     #
#     #     st.text("")
#     #     st.text("")
#     #     st.text("")
#     #
#     #     _k1, kpi1, kpi2, _k_2 = st.columns(4)
#     #
#     #     # with kpi1:
#     #     revenue_2024 = select_Site_data[" Revenue"].sum() / 1000
#     #     # revenue_2023 = select_Site_data_2023_profitability[" Revenue"].sum() / 1000
#     #     # delta = f"{((revenue_2024 / revenue_2023) - 1):,.1%}"
#     #
#     #     kpi1.metric(label=f"FY24: Revenue - 000s", value=f"{revenue_2024:,.0f}")
#     #
#     #     # with kpi2:
#     #     ebitda_2024 = select_Site_data["EBITDA $"].sum() / 1000
#     #     # ebitda_2023 = select_Site_data_2023_profitability["EBITDA $"].sum() / 1000
#     #     # delta = f"{((ebitda_2024 / ebitda_2023) - 1):,.1%}"
#     #
#     #     ebitda_margin = f"{(ebitda_2024 / revenue_2024):,.1%}"
#     #
#     #     kpi2.metric(label=f"FY24: EBITDA - 000s", value=f"{ebitda_2024:,.0f}",
#     #                 delta=f"| FY24 → {ebitda_margin}   ")
#     #     st.text("")
#     #     st.text("")
#     #
#     #     _k_3, kpi3, kpi4, _k_4 = st.columns(4)
#
#     # dominic_report_2024 = dominic_report[
#     #     (dominic_report.Year == 2024) & (dominic_report.Month == "December")]
#     # dominic_report_2024["Site"] = dominic_report_2024["Cost Centers"].apply(extract_site)
#     # dominic_report_2024 = dominic_report_2024[dominic_report_2024.Site.isin(selected_site)]
#
#     # dominic_report_2023 = dominic_report[
#     #     (dominic_report.Year == 2023) & (dominic_report.Month == "December")]
#     # dominic_report_2023["Site"] = dominic_report_2023["Cost Centers"].apply(extract_site)
#     # dominic_report_2023 = dominic_report_2023[dominic_report_2023.Site.isin(selected_site)]
#     #
#     # economic_pallets_2024 = dominic_report_2024["Economic Pallet Total"].sum()
#     # capacity_2024 = dominic_report_2024.Denominator.sum()
#
#     # economic_pallets_2023 = dominic_report_2023["Economic Pallet Total"].sum()
#     # capacity_2023 = dominic_report_2023.Denominator.sum()
#     #
#     # dec_2024_occupancy = economic_pallets_2024 / capacity_2024
#     # dec_2023_occupancy = economic_pallets_2023 / capacity_2023
#     #
#     # delta = f"{((dec_2024_occupancy / dec_2023_occupancy) - 1):,.1%}"
#     #
#     # kpi3.metric(label=f"Economic Occupancy", value=f"{dec_2024_occupancy:,.1%}", delta=f"{delta} ")
#     #
#     # # with kpi4:
#     # volume_guarantee_2024 = select_Site_data_2023_pallets["VG Pallets - Dec-2024"].sum()
#     # volume_guarantee_2023 = select_Site_data_2023_pallets["VG Pallets - Dec-2023"].sum()
#     #
#     # delta = f"{(volume_guarantee_2024 / capacity_2024):,.1%} : of Capacity is VG "
#     #
#     # kpi4.metric(label=f"Volume Guarantee", value=f"{volume_guarantee_2024:,.0f}", delta=f"{delta} ")
#     # st.text("")
#     # st.text("")