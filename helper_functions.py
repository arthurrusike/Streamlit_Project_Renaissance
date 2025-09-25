from altair.theme import active

import streamlit as st
import pandas as pd
import pyodbc


def style_dataframe(df):
    return (
        df.set_table_styles(

            [
                {
                'selector': 'th',
                'props': [
                    ('background-color', '#305496'),
                    ('width', 'auto'),
                    ('color', 'white'),
                    ('font-family', 'sans-serif, Arial'),
                    ('font-size', '12px'),
                    ('text-align', 'center')
                ]
            },
                {
                    'selector': ' th',
                    'props': [
                        ('border', '2px solid white')
                    ]
                },
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
        )

    )


def extract_cc(cost_centre):
    return cost_centre[:9]


def extract_name(value):
    return value.split(" [")[0].title()


def sub_category_classification(dtframe):
    Storage_Revenue = ['Storage - Initial', 'Storage - Renewal', 'Storage Guarantee']
    Handling_Revenue = ['Handling - Initial', 'Handling Out']
    Case_Pick_Revenue = ['Accessorial - Case Pick / Sorting']
    Ancillary_Revenue = ['Accessorial - Documentation', 'Accessorial - Labeling / Stamping',
                         'Accessorial - Labor and Overtime', 'Accessorial - Loading / Unloading / Lumping',
                         'Accessorial - Palletizing', 'Accessorial - Shrink Wrap']
    Blast_Freezing_Revenue = ['Blast Freezing', 'Room Freezing']
    Other_Warehouse_Revenue = ['Other - Delayed Pallet Hire Revenue', 'Other - Warehouse Revenue',
                               'Rental Electricity Income']

    if dtframe["Revenue_Category"] in Storage_Revenue:
        return "Storage Revenue"
    elif dtframe["Revenue_Category"] in Handling_Revenue:
        return "Handling Revenue"
    elif dtframe["Revenue_Category"] in Case_Pick_Revenue:
        return "Case Pick Revenue"
    elif dtframe["Revenue_Category"] in Ancillary_Revenue:
        return "Ancillary Revenue"
    elif dtframe["Revenue_Category"] in Blast_Freezing_Revenue:
        return "Blast Freezing Revenue"
    elif dtframe["Revenue_Category"] in Other_Warehouse_Revenue:
        return "Other Warehouse Revenue"
    else:
        return dtframe["Revenue_Category"]


@st.cache_data
def load_profitbility_Summary_model(uploaded_file):
    customer_profitability_summary = pd.read_excel(uploaded_file, sheet_name="ChartData", header=5, usecols="B:BR")
    customer_profitability_summary = customer_profitability_summary[customer_profitability_summary[" Revenue"] > 0]
    return customer_profitability_summary


@st.cache_data
def load_rates_standardisation(uploaded_file):
    excel_file = pd.ExcelFile(uploaded_file)
    return excel_file


@st.cache_data
def load_specific_xls_sheet(file, sheet_name, header, use_cols):
    cached_xls_sheet = pd.read_excel(file, sheet_name=sheet_name, header=header, usecols=use_cols)
    return cached_xls_sheet

@st.cache_data
def run_sql_query(startDate, endDate):

    # conn = pyodbc.connect('DSN=CalumoCoreDW; Trusted_Connection=yes; DRIVER=SQL Server;')
    # conn = pyodbc.connect('DSN=CalumoCoreDW; Trusted_Connection=yes; DRIVER=SQL Server;')
    # conn = pyodbc.connect('Driver={{SQL Server}};Server=au-cl1-dwdb\dwprod;Database= CoreDW;Trusted_Connection=yes;DSN=CalumoCoreDW;')
    conn = pyodbc.connect('Trusted_Connection=yes;DSN=CalumoCoreDW;')

    slqQuery = f"SELECT * from CoreDW.[stgAQT].[vwRates] where [InvoiceDate] between '{str(startDate)}' AND '{str(endDate)}' and [Cost_Center] like '%S&H%' Order By SourceSystem, InvoiceNumber"
    invoice_rates = pd.read_sql_query(slqQuery, conn)

    return invoice_rates



def style_commodity_customers(df):
    return (
        df.set_table_styles(

            [
                {
                    'selector': 'th',
                    'props': [
                        ('background-color', '#305496'),
                        ('width', 'auto'),
                        ('color', 'white'),
                        ('font-family', 'sans-serif, Arial'),
                        ('font-size', '12px'),
                        ('text-align', 'center')
                    ]
                },
                {
                    'selector': ' th',
                    'props': [
                        ('border', '2px solid white')
                    ]
                },
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
                        ('text-align', 'left'),
                        ('font-size', '12px'),
                        ('font-family', 'sans-serif, Arial'),

                    ]
                }
            ]
        )

    )