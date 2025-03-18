# import io
# import random
# from io import BytesIO
# import numpy as np
# import pandas as pd
# import plotly.express as px
# import plotly.graph_objects as go
# import streamlit as st
#
# from site_detailed_view_helper_functions import load_invoices_model, profitability_model
#
# st.set_page_config(page_title="Customer Rates and Invoicing Analysis", page_icon="📊", layout="wide",
#                    initial_sidebar_state="expanded")
#
# ############################### Side for Uploading excel File ###############################
#
# with st.spinner('Loading and Updating Report...🥱'):
#     try:
#         with st.sidebar:
#             st.title("Analytics Dashboard")
#             uploaded_file = st.file_uploader('Upload Customer Rates File', type=['xlsx', 'xlsm', 'xlsb'])
#             # Main Dashboard
#             invoice_rates = load_invoices_model(uploaded_file)
#             st.divider()
#
#             st.markdown("#### Dates Covered by Invoice File")
#             start_date = invoice_rates.formatted_date.dropna()
#             st.markdown(f'###### Report Start Date  : {start_date.min()}')
#             st.markdown(f'###### Report End Date  : {start_date.max()}')
#             st.divider()
#
#             st.title("Profitability")
#             st.markdown(f'###### This is the Customer Pricing Model from Box:APAC Rev...')
#             profitability_model_file = st.file_uploader('Upload Customer Pricing Model', type=['xlsx', 'xlsm', 'xlsb'])
#
#             st.divider()
#
#
#     except Exception as e:
#         st.divider()
#         st.subheader('Upload Customer Rates File to use with this App 📊 ')
#         print(e)
#
# ###############################  Main Dashboard ##############################
#
# if uploaded_file:
#     cost_centres = invoice_rates[invoice_rates.Cost_Center.str.contains("S&H", na=False)]
#     cost_centres = cost_centres.Cost_Center.unique()
#     all_workday_Customer_names = invoice_rates.WorkdayCustomer_Name.unique()
#
#     select_Option1, select_Option2 = st.columns(2)
#
#     with select_Option1:
#         selected_cost_centre = st.selectbox("Cost Centre :", cost_centres, index=0)
#         selected_workday_customers = invoice_rates[
#             invoice_rates.Cost_Center == selected_cost_centre].WorkdayCustomer_Name.unique()
#         select_CC_data = invoice_rates[invoice_rates.Cost_Center == selected_cost_centre]
#
#     with select_Option2:
#         if selected_cost_centre:
#             select_customer = st.selectbox("Select Customer : ", selected_workday_customers, index=0)
#         else:
#             select_customer = st.selectbox("Select Customer : ", all_workday_Customer_names, index=0)
#
#     st.divider()
#
#     DetailsTab, SiteTab, ProfitabilityTab, AboutTab = st.tabs(["📊 Quick View", "🥇 Site View",
#                                                                "📈 Estimate Profitability",
#                                                                "ℹ️ About"])
#
#     ###############################  Details Tab ###############################
#     with DetailsTab:
#
#         Workday_Sales_Item_Name = invoice_rates.Workday_Sales_Item_Name.unique()
#         selected_customer = select_CC_data[select_CC_data.WorkdayCustomer_Name == select_customer]
#
#         Revenue_Category = selected_customer.Revenue_Category.dropna().unique()
#
#         selected_customer_pivot = pd.pivot_table(selected_customer,
#                                                  values=["Quantity", "LineAmount"],
#                                                  index=["Revenue_Category", "Workday_Sales_Item_Name"],
#                                                  aggfunc="sum")
#         selected_customer_pivot["Avg Rate"] = selected_customer_pivot.LineAmount / selected_customer_pivot.Quantity
#
#         col1, col2 = st.columns(2)
#
#         with col1:
#             customer_avg_plts = selected_customer[selected_customer['Revenue_Category'] == 'Storage - Renewal'][
#                 "Quantity"].mean()
#             site_avg_plts_occupied = select_CC_data[select_CC_data['Revenue_Category'] == 'Storage - Renewal'][
#                 "Quantity"].mean()
#
#             st.dataframe(selected_customer_pivot, use_container_width=True)
#
#             # Convert DataFrame to Excel
#             output = io.BytesIO()
#             with pd.ExcelWriter(output) as writer:
#                 selected_customer_pivot.to_excel(writer)
#
#             # Create a download button
#             st.download_button(
#                 label="Download Excel  ⤵️",
#                 data=output,
#                 file_name='invoiced_data.xlsx',
#                 mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
#                 key=1
#             )
#
#             kpi1, kpi2, kpi3, = st.columns(3)
#
#             with kpi1:
#                 # delta_value = (customer_avg_plts / site_avg_plts_occupied)/100
#                 # formartted_value = "{:.2%}".format(delta_value)
#                 st.metric(label=" Customer Avg Recurring Storage", value=f"{customer_avg_plts:,.0f} Plts",
#                           # delta=f"~{formartted_value} of CC Holding",
#                           border=True)
#
#             with kpi2:
#                 pallet_billed_services = ['Accessorial - Shrink Wrap', 'Blast Freezing', 'Handling - Initial',
#                                           'Handling Out',
#                                           'Storage - Initial', 'Storage - Renewal']
#
#                 filtered_data = selected_customer[selected_customer["Revenue_Category"].isin(pallet_billed_services)][
#                     "LineAmount"].sum()
#                 ancillary_data = selected_customer["LineAmount"].sum() - filtered_data
#
#                 st.metric(label="Billed Pallet Rev  ", value=f"${filtered_data:,.0f} ", delta="excl Volume Guarantee ",
#                           border=True)
#
#             with kpi3:
#                 st.metric(label="Billed Ancillary Rev", value=f"${ancillary_data:,.0f}",
#                           border=True, )
#
#         with col2:
#             selected_graph = st.selectbox("Display :", ["Revenue", "Volumes"], index=0)
#             graph_values = selected_graph
#
#             if selected_graph == "Revenue":
#                 graph_values = "LineAmount"
#             else:
#                 graph_values = "Quantity"
#
#             fig = px.pie(selected_customer_pivot, values=selected_customer_pivot[graph_values],
#                          names=selected_customer_pivot.index.get_level_values(0),
#                          title=f'{selected_graph} View',
#                          height=300, width=200)
#             fig.update_layout(margin=dict(l=20, r=20, t=30, b=0), )
#             st.plotly_chart(fig, use_container_width=True)
#
#         st.divider()
#
#         ###############################  Pallet Positions vs Thru Put Pallets  ###############################
#
#         select_CC_data_tab1 = select_CC_data[select_CC_data.WorkdayCustomer_Name == select_customer]
#         storage_holding = select_CC_data_tab1[select_CC_data_tab1['Revenue_Category'] == 'Storage - Renewal']
#         storage_holding.set_index("InvoiceDate", inplace=True)
#         storage_holding = storage_holding.Quantity
#
#         pallet_Thru_Put = select_CC_data_tab1[(select_CC_data_tab1.Revenue_Category == 'Handling - Initial') |
#                                               (select_CC_data_tab1.Revenue_Category == 'Handling Out')]
#
#         pallet_Thru_Put_pivot = pd.pivot_table(pallet_Thru_Put,
#                                                values=["Quantity"],
#                                                index=["InvoiceDate"],
#                                                aggfunc="sum")
#
#         fig2 = go.Figure()
#         fig2.add_trace(
#             go.Bar(x=pallet_Thru_Put_pivot.index, y=storage_holding, name="Pallet Holding"))
#
#         fig2.add_trace(go.Scatter(x=pallet_Thru_Put_pivot.index, y=pallet_Thru_Put_pivot.Quantity, mode="lines",
#                                   name="Plt Thru Put", yaxis='y2'))
#         fig2.update_layout(
#             title="Pallet Position vs Thru Put Pallets",
#             xaxis=dict(title=select_customer),
#             yaxis=dict(title="Pallets Stored", showgrid=False),
#             yaxis2=dict(title="Pallet Thru Put", overlaying='y', side="right"),
#             template="gridon",
#             legend=dict(x=1, y=1)
#
#         )
#
#         st.plotly_chart(fig2, use_container_width=True)
#
#         st.divider()
#         sdiv, sdiv2, sdiv3 = st.columns((1, 4, 1))
#         sdiv.write("")
#
#         ###############################  Box Plots Over and under Indexed Customers  ###############################
#
#         sdiv2.subheader('Outliers View:- Under vs Over Indexed Customer Rates')
#         sdiv3.write("")
#         st.divider()
#
#         p1_Storage, p2_Handling, p3_Wrapping = st.columns((1, 1, 1))
#
#         with p1_Storage:
#             df_pStorage = \
#                 select_CC_data[(select_CC_data['Revenue_Category'] == 'Storage - Initial') |
#                                (select_CC_data['Revenue_Category'] == 'Storage - Renewal')
#                                ].groupby(
#                     "WorkdayCustomer_Name")[
#                     "UnitPrice"].mean()
#
#             customer_names = df_pStorage.index
#             q3_pStorage = np.percentile(df_pStorage.values, 75)
#             q1_pStorage = np.percentile(df_pStorage.values, 25)
#
#             # #### Storage Box plot ####################################################################
#             # # Add points with Labels
#             y = df_pStorage.values
#             x = np.random.normal(1, 0.250, size=(len(y)))
#
#             # Create Box Plot Figure:
#             box_x_value = np.random.normal(1, 0.250, size=len(y))
#
#             # Create a Boxplot
#             fig = go.Figure()
#
#             fig.add_trace(go.Box(
#                 y=y,
#                 boxpoints='all',
#                 jitter=0.3,
#                 pointpos=0.5,
#                 name="Storage Rate",
#             ))
#
#             # Add annotations for each data point. to recheck
#             for i, (val, customer) in enumerate(zip(y, customer_names)):
#                 fig.add_annotation(
#                     y=val,
#                     showarrow=False,
#                     xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
#                     hovertext=f'{customer}, Rate:  {val}',
#                     text="*",
#
#                 )
#                 if customer == select_customer:
#                     customer = select_customer
#                     fig.update_annotations(
#                         y=val,
#                         text=f'{customer}, Rate:  {val}',
#                         showarrow=False,
#                         xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
#                         hovertext=f'{customer}, Rate:  {val}',
#                         # name=f'Customer: {customer}, Rate:  {val}',
#                     )
#
#             fig.update_layout(
#                 title=dict(text='Cost Centre Storage Rates'),
#                 margin=dict(l=0, r=10, b=10, t=30),
#             )
#
#             p1_Storage.plotly_chart(fig)
#
#         with p2_Handling:
#
#             #### Handling Box plot ####################################################################
#             df_pHandling = select_CC_data[(select_CC_data['Revenue_Category'] == 'Handling - Initial') &
#                                           (select_CC_data['Revenue_Category'] == 'Handling - Initial')].groupby(
#                 "WorkdayCustomer_Name")["UnitPrice"].mean()
#
#             handling_customer_names = df_pHandling.index
#             q3_pHandling = np.percentile(df_pHandling.values, 75)
#             q1_pHandling = np.percentile(df_pHandling.values, 25)
#
#             # Add points with Labels
#             y = df_pHandling.values
#             x = np.random.normal(1, 0.250, size=(len(y)))
#
#             # Create Box Plot Figure:
#             box_x_value = np.random.normal(1, 0.250, size=len(y))
#
#             # Create a Boxplot
#             fig = go.Figure()
#
#             fig.add_trace(go.Box(
#                 y=y,
#                 boxpoints='all',
#                 jitter=0.3,
#                 pointpos=0.5,
#                 name="Handling Rates"
#             ))
#
#             # Add annotations for each data point
#             for i, (val, customer) in enumerate(zip(y, handling_customer_names)):
#                 fig.add_annotation(
#                     y=val,
#                     text="*",
#                     showarrow=False,
#                     xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
#                     hovertext=f'{customer}, Rate:  {val}',
#                 )
#                 if customer == select_customer:
#                     customer = select_customer
#                     fig.update_annotations(
#                         y=val,
#                         text=f'{customer}, Rate:  {val}',
#                         showarrow=False,
#                         xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
#                         hovertext=f'{customer}, Rate:  {val}',
#                         # name=f'Customer: {customer}, Rate:  {val}',
#                     )
#
#             fig.update_layout(
#                 title=dict(text='Spread of Customer Handling Rates'),
#                 margin=dict(l=0, r=10, b=10, t=30)
#             )
#
#             p2_Handling.plotly_chart(fig)
#
#         try:
#
#             with p3_Wrapping:
#                 ############################ Shrink Wrap Box plot ######################################################
#                 df_pWrapping = select_CC_data[select_CC_data['Revenue_Category'] == 'Accessorial - Shrink Wrap'] \
#                     .groupby("WorkdayCustomer_Name")["UnitPrice"].mean()
#
#                 wrapping_customer_names = df_pWrapping.index
#                 q3_pWrapping = np.percentile(df_pWrapping.values, 75)
#                 q1_pWrapping = np.percentile(df_pWrapping.values, 25)
#
#                 # Add points with Labels
#                 y = df_pWrapping.values
#                 x = np.random.normal(1, 0.250, size=(len(y)))
#
#                 # Create Box Plot Figure:
#                 box_x_value = np.random.normal(1, 0.250, size=len(y))
#
#                 # Create a Boxplot
#                 fig = go.Figure()
#                 fig.add_trace(go.Box(
#                     y=y,
#                     boxpoints='all',
#                     jitter=0.3,
#                     pointpos=0.5,
#                     name="Shrink Wrap Rates"
#                 ))
#
#                 # Add annotations for each data point
#                 for i, (val, customer) in enumerate(zip(y, wrapping_customer_names)):
#                     fig.add_annotation(
#                         y=val,
#                         showarrow=False,
#                         xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
#                         hovertext=f'{customer}, Rate:  {val}',
#                         text="*",
#                     )
#                     if customer == select_customer:
#                         customer = select_customer
#                         fig.update_annotations(
#                             y=val,
#                             text=f'{customer}, Rate:  {val}',
#                             showarrow=False,
#                             xshift=x[i] * -200,  # Adjust this value to position the labels horizontally
#                             hovertext=f'{customer}, Rate:  {val}',
#                             # name=f'Customer: {customer}, Rate:  {val}',
#                         )
#
#                 fig.update_layout(
#                     title=dict(text='Spread of Customer Pallet Wrapping Rates'),
#                     margin=dict(l=0, r=10, b=10, t=20))
#
#                 p3_Wrapping.plotly_chart(fig)
#         except Exception as e:
#             st.markdown(f"##############  Data pulled for this Customer or Site does not have all the Columns needed "
#                         f"to display info")
#             pass
#
#     with SiteTab:
#         selected_cost_Centre = select_CC_data[
#             ["WorkdayCustomer_Name", "Revenue_Category", "LineAmount", "Quantity"]].fillna(0)
#
#         selected_cost_Centre = selected_cost_Centre.groupby(["WorkdayCustomer_Name", "Revenue_Category"]
#                                                             ).agg(
#             {"LineAmount": 'sum', "Quantity": 'sum', }).reset_index()
#
#         selected_LineAmount = selected_cost_Centre.pivot(index="WorkdayCustomer_Name",
#                                                          columns="Revenue_Category",
#                                                          values="LineAmount").reset_index()
#
#         selected_LineAmount = selected_LineAmount.fillna(0)
#
#         selected_LineAmount["Total"] = selected_LineAmount.iloc[:, 1:].sum(axis=1)
#
#         selected_LineAmount = selected_LineAmount.nlargest(10, ["Total"])
#
#         selected_LineAmount.reindex()
#
#         columns = selected_LineAmount.columns
#         colours = ['#7fbf7b',
#                    '#9FE2BF',
#                    '#DE3163',
#                    '#2b83ba',
#                    '#6495ED',
#                    '#FFBF00',
#                    '#40E0D0',
#                    "#CCCCFF",
#                    '#FF7F50']
#
#         marker_color_map = {}
#
#         for i in range(len(columns)):
#             if columns[i] == 'WorkdayCustomer_Name' or columns[i] == 'Total':
#                 pass
#             else:
#                 random_colour = {columns[i]: colours[random.randint(0, 8)]}
#                 marker_color_map |= random_colour
#
#         #### Plotting HEAT MAP - tow Show Revenue and Profitabiliy #########################################
#
#         st.markdown('##### Avg Recurring Pallet Storage Billed')
#         st.divider()
#
#         selected_customer_vol_billed = pd.pivot_table(select_CC_data,
#                                                       values=["Quantity"],
#                                                       index=["Revenue_Category", "WorkdayCustomer_Name"],
#                                                       aggfunc="mean")
#
#         selected_customer_vol_billed = selected_customer_vol_billed.query(
#             "WorkdayCustomer_Name != '.All Other [.All Other]'")
#         treemap_data = selected_customer_vol_billed.loc["Storage - Renewal"]
#         customers = treemap_data.index
#
#         fig = px.treemap(treemap_data,
#                          path=[customers],
#                          values=treemap_data["Quantity"],
#                          color=treemap_data["Quantity"],
#                          custom_data=[customers],
#                          color_discrete_map=colours,
#                          # title="Testing",
#                          )
#
#         st.plotly_chart(fig, filename='Chart.html')
#         st.divider()
#
#         selected_LineAmount.rename(
#             columns={0: "Other"},
#             inplace=True)
#
#         sortedListed = selected_LineAmount.columns.sort_values(ascending=False)
#         download_data = selected_LineAmount[list(sortedListed)]
#         st.dataframe(download_data, hide_index=True)
#
#         # Convert DataFrame to Excel
#
#         output2 = io.BytesIO()
#         with pd.ExcelWriter(output2 ) as writer:
#             download_data.to_excel(writer, sheet_name='export_data', index=False)
#
#
#         # Create a download button
#         st.download_button(
#             label="Download file️ ⤵️",
#             data=output2,
#             file_name='customer_revenue.xlsx',
#             mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
#             key=2,
#         )
#
#     st.divider()
#     try:
#         with ProfitabilityTab:
#
#             AU_facility_financials = profitability_model(profitability_model_file)
#             selected_cost_Centre_financials = AU_facility_financials[
#                 AU_facility_financials["CC Long #"] == selected_cost_centre[0:9]]
#             selected_cost_centre_row_number = list(selected_cost_Centre_financials.index)[0]
#             revenue_amount = selected_cost_Centre_financials["Revenue"]
#             selected_cost_Centre_financials = selected_cost_Centre_financials[selected_cost_Centre_financials.columns]
#             total_labour = selected_cost_Centre_financials["Total Labor and Benefits"]
#             operating_expenses = selected_cost_Centre_financials["Total Operating Expenses"]
#
#             selected_customer_pivot_profitability = pd.pivot_table(selected_customer,
#                                                                    values=["LineAmount"],
#                                                                    index=["Calumo Description"],
#                                                                    aggfunc="sum")
#             selected_customer_pivot_profitability["% Share"] = [f"{i / (selected_customer["LineAmount"].sum()):.1%}" for
#                                                                 i in
#                                                                 selected_customer_pivot_profitability["LineAmount"]]
#
#             selected_customer_pivot_profitability = selected_customer_pivot_profitability.style.format(
#                 {'LineAmount': lambda x: '{:,.0f}'.format(x)})
#
#             prof1, prof2 = st.columns(2)
#
#             with prof1:
#                 st.write(f""" ##### {select_customer}  Revenues """)
#                 st.dataframe(selected_customer_pivot_profitability)
#
#             with prof2:
#                 revenue_lines = selected_cost_Centre_financials[selected_cost_Centre_financials.columns[3:22]].T
#                 revenue_lines = revenue_lines.reset_index()
#                 revenue_lines.rename(
#                     columns={"index": "Revenue Categories", selected_cost_centre_row_number: "Amount"},
#                     inplace=True)
#                 revenue_lines = revenue_lines[revenue_lines.Amount > 0]
#                 revenue_lines["% Share"] = [f"{i / revenue_amount[selected_cost_centre_row_number]:.1%}" for i in
#                                             revenue_lines["Amount"]]
#                 revenue_lines = revenue_lines.style.format(
#                     {'Amount': lambda x: '{:,.0f}'.format(x)})
#
#                 st.write(f""" ##### {selected_cost_centre} : Revenue """)
#                 st.dataframe(revenue_lines, hide_index=True)
#
#                 st.write(""" #####   Labour & Benefits """)
#                 exp = st.expander(f"Total Labour and Benefits............:\t {list(total_labour)[0]:,.0f}")
#                 total_labour_dataframe = selected_cost_Centre_financials[
#                     selected_cost_Centre_financials.columns[22:47]].T
#                 total_labour_dataframe = total_labour_dataframe.reset_index()
#                 total_labour_dataframe = total_labour_dataframe.rename(
#                     columns={"index": "Revenue Categories", selected_cost_centre_row_number: "Amount"})
#                 total_labour_dataframe = total_labour_dataframe[total_labour_dataframe["Amount"] > 0]
#                 total_labour_dataframe = total_labour_dataframe.style.format(
#                     {'Amount': lambda x: '{:,.0f}'.format(x)})
#                 exp.dataframe(total_labour_dataframe, hide_index=True)
#
#                 st.write(""" #####   Operating Expenses """)
#                 exp = st.expander(f"Operating Expenses.....................: \t {list(operating_expenses)[0]:,.0f}")
#                 operating_expenses_dataframe = selected_cost_Centre_financials[
#                     selected_cost_Centre_financials.columns[47:71]].T
#                 operating_expenses_dataframe.reset_index(inplace=True)
#                 operating_expenses_dataframe.rename(
#                     columns={"index": "Revenue Categories", selected_cost_centre_row_number: "Amount"},
#                     inplace=True)
#                 operating_expenses_dataframe = operating_expenses_dataframe[operating_expenses_dataframe.Amount > 0]
#                 operating_expenses_dataframe = operating_expenses_dataframe.style.format(
#                     {'Amount': lambda x: '{:,.0f}'.format(x)})
#
#                 exp.dataframe(operating_expenses_dataframe, hide_index=True)
#
#                 ebitda_details1, ebitda_details2, ebitda_details3, ebitda_details4, = st.columns(4)
#
#                 site_capacity = list(selected_cost_Centre_financials["Total Pallet Capacity"])[0]
#                 site_sqm = list(selected_cost_Centre_financials["Square Meters"])[0]
#                 total_expenses = list(selected_cost_Centre_financials["Total Expenses"])[0]
#                 ebitdar_amount = list(selected_cost_Centre_financials["EBITDAR"])[0]
#                 rent_amount = list(selected_cost_Centre_financials["700100:Rent Expense (Real Estate)"])[0]
#                 ebitda_amount = ebitdar_amount - rent_amount
#
#                 ebitda_details1.text(
#                     f"Total Expenses :\n {'{:,.0f}'.format(total_expenses)}")
#                 ebitda_details2.text(f"EBITDAR :\n {'{:,.0f}'.format(ebitdar_amount)}")
#                 ebitda_details3.text(f"RENT :\n {'{:,.0f}'.format(rent_amount)}")
#                 ebitda_details4.text(f"EBITDA :\n {'{:,.0f}'.format(ebitda_amount)}")
#
#                 st.divider()
#
#                 stats1, stats2, stats3, stats4 = st.columns(4)
#
#                 stats1.text(f"Site SQM: \n {'{:,.0f}'.format(site_sqm)}")
#                 stats2.text(
#                     f"EBITDAR/m²: \n {'{:,.0f}'.format(ebitdar_amount / site_sqm)}")
#                 stats3.text(
#                     f"Rev/m²: \n {'{:,.0f}'.format(revenue_amount[selected_cost_centre_row_number] / site_sqm)}")
#                 stats4.text(f"EBITDA/m²: \n {'{:,.0f}'.format(ebitda_amount / site_sqm)}")
#     except Exception as e:
#         st.markdown("###### Please Check if Customer Pricing Model has been Uploaded")
#
#     with AboutTab:
#         st.markdown(f'Check in with Arthur Rusike')
