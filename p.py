import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

# Title and description of the app
st.title("P2P Analysis")
st.write("Upload your Excel file to analyze P2P data and explore interactive visualizations.")

# File uploader widget
uploaded_file = st.file_uploader("Drag and drop your Excel file here", type=["xlsx"])

# Initialize session state flag for processed data
if 'processed' not in st.session_state:
    st.session_state.processed = False

# Process the file if uploaded and not yet processed
if uploaded_file is not None and not st.session_state.processed:
    # Read and process the data
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    
    # Convert date columns to datetime format
    df['Document Date'] = pd.to_datetime(df['Document Date'], errors='coerce')
    df['Delivery Date'] = pd.to_datetime(df['Delivery Date'], errors='coerce')

    # Create Month column early to ensure availability
    df['Month'] = df['Document Date'].dt.to_period('M').astype(str)

    # Check for Delivery Date anomalies
    df['Delivery_Date_Anomaly'] = df['Delivery Date'] < df['Document Date']
    df['Date_Difference'] = (df['Delivery Date'] - df['Document Date']).dt.days
    date_anomalies = df[df['Delivery_Date_Anomaly']]

    # Calculate delivery delay and flag backdating issues
    df['Delivery Delay'] = (df['Delivery Date'] - df['Document Date']).dt.days.fillna(-1)
    df['Why PO Delay'] = df['Delivery Delay'].apply(lambda x: 'PO raised after delivery' if x < 0 else '')

    # Ensure numeric columns are correctly converted
    numeric_cols = ['PO Ordered Value in Loc. Curr.', 'PO Invoice Value in Loc. Curr.', 
                    'Ordered Quantity', 'Delivery Quantity', 'PO Down Payment', 'Still to Deliver']
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce')

    # Identify overbilling cases (invoice > order)
    df['Overbilling_Flag'] = df['PO Invoice Value in Loc. Curr.'] > df['PO Ordered Value in Loc. Curr.']
    overbilling_cases = df[df['Overbilling_Flag']]

    # Calculate the overbilling amount (invoice - order)
    df['Overbilling Amount'] = df['PO Invoice Value in Loc. Curr.'] - df['PO Ordered Value in Loc. Curr.']

    # For focused overbilling analysis, filter records with positive differences
    overbilling_df = df[df['Overbilling Amount'] > 0]

    # Define current date and check for on-time delivery
    current_date = pd.Timestamp.today()
    df['On_Time'] = (df['Delivery Date'] >= current_date) | (df['Still to Deliver'] == 0)

    # 1. Total Spend by Vendor
    total_spend_by_vendor = df.groupby(['Vendor Name', 'Vendor Number', 'Entity Name', 'IT/NON-IT']).agg({
        'PO Ordered Value in Loc. Curr.': 'sum',
        'PO Invoice Value in Loc. Curr.': 'sum',
        'PO Down Payment': 'sum'
    }).reset_index()
    total_spend_by_vendor.columns = ['Vendor Name', 'Vendor Number', 'Entity Name', 'IT/NON-IT', 
                                      'Total PO Ordered Value', 'Total PO Invoice Value', 'Total PO Down Payment']

    # 2. Total Spend by Material
    total_spend_by_material = df.groupby(['Material Description', 'IT/NON-IT']).agg({
        'PO Ordered Value in Loc. Curr.': 'sum',
        'PO Invoice Value in Loc. Curr.': 'sum',
        'PO Down Payment': 'sum'
    }).reset_index()
    total_spend_by_material.columns = ['Material Description', 'IT/NON-IT', 
                                        'Total PO Ordered Value', 'Total PO Invoice Value', 'Total PO Down Payment']

    # 3. Total Spend by Service Area
    total_spend_by_service_area = df.groupby(['Service Area', 'IT/NON-IT']).agg({
        'PO Ordered Value in Loc. Curr.': 'sum',
        'PO Invoice Value in Loc. Curr.': 'sum',
        'PO Down Payment': 'sum'
    }).reset_index()
    total_spend_by_service_area.columns = ['Service Area', 'IT/NON-IT', 
                                            'Total PO Ordered Value', 'Total PO Invoice Value', 'Total PO Down Payment']

    # 4. Top 10 Vendors by Spend (using Total PO Ordered Value)
    top_10_vendors = total_spend_by_vendor.sort_values(by='Total PO Ordered Value', ascending=False).head(10)

    # 5. Top 10 Materials by Spend (using Total PO Ordered Value)
    top_10_materials = total_spend_by_material.sort_values(by='Total PO Ordered Value', ascending=False).head(10)

    # 6. Spend Trends Over Time (Monthly)
    monthly_spend = df.groupby(['Month', 'Vendor Name', 'Vendor Number', 'Entity Name', 'IT/NON-IT']).agg({
        'PO Ordered Value in Loc. Curr.': 'sum',
        'PO Invoice Value in Loc. Curr.': 'sum',
        'PO Down Payment': 'sum'
    }).reset_index()
    monthly_spend = monthly_spend.sort_values(by=['Month', 'PO Ordered Value in Loc. Curr.'], ascending=[True, False])
    top_10_vendors_monthly = monthly_spend.groupby('Month').head(10).reset_index(drop=True)
    top_10_vendors_monthly.columns = ['Month', 'Vendor Name', 'Vendor Number', 'Entity Name', 'IT/NON-IT', 
                                      'Total PO Ordered Value', 'Total PO Invoice Value', 'Total PO Down Payment']

    # 7. Vendor Order Summary with Delivery Percentage
    vendor_summary = df.groupby(['Vendor Name', 'Vendor Number', 'Entity Name', 'IT/NON-IT']).agg(
        Total_Orders=('Ordered Quantity', 'sum'),
        Total_Delivered=('Delivery Quantity', 'sum'),
        Total_Pending=('Still to Deliver', 'sum'),
        PO_Ordered_Value=('PO Ordered Value in Loc. Curr.', 'sum'),
        PO_Invoice_Value=('PO Invoice Value in Loc. Curr.', 'sum'),
        PO_Down_Payment=('PO Down Payment', 'sum')
    ).reset_index()
    vendor_summary['Delivery_Percentage'] = np.where(
        vendor_summary['Total_Orders'] > 0,
        (vendor_summary['Total_Delivered'] / vendor_summary['Total_Orders']) * 100,
        0
    ).round(2)
    vendor_summary.columns = [
        'Vendor Name', 'Vendor Number', 'Entity Name', 'IT/NON-IT',
        'Total Ordered Quantity', 'Total Delivered Quantity',
        'Total Pending Quantity', 'Total PO Ordered Value',
        'Total PO Invoice Value', 'Total PO Down Payment',
        'Delivery Percentage (%)'
    ]

    # 8. Delayed POs
    delayed_pos = df[(df['GR Document Number'].isna()) | (df['IR Document Number'].isna())]
    output_columns = [
        'Purchasing Document Number', 'Document Date', 'Delivery Date',
        'Delivery Delay', 'Why PO Delay', 'IT/NON-IT', 'Vendor Number', 'Entity Name'
    ]
    existing_columns = [col for col in output_columns if col in df.columns]
    delayed_pos = delayed_pos[existing_columns]

    # 9. Quantity Errors
    quantity_errors = df[df['Delivery Quantity'] > df['Ordered Quantity']]

    # New visualizations data preparation
    spend_trend = df.groupby('Month').agg({
        'PO Ordered Value in Loc. Curr.': 'sum',
        'PO Invoice Value in Loc. Curr.': 'sum'
    }).reset_index()

    vendor_spend = df.groupby('Vendor Name').agg({
        'PO Ordered Value in Loc. Curr.': 'sum',
        'PO Invoice Value in Loc. Curr.': 'sum'
    }).reset_index()

    entity_spend = df.groupby('Entity Name')['PO Ordered Value in Loc. Curr.'].sum().reset_index()

    # Store all processed data in session state
    st.session_state.update({
        'total_spend_by_vendor': total_spend_by_vendor,
        'total_spend_by_material': total_spend_by_material,
        'total_spend_by_service_area': total_spend_by_service_area,
        'top_10_vendors': top_10_vendors,
        'top_10_materials': top_10_materials,
        'top_10_vendors_monthly': top_10_vendors_monthly,
        'vendor_summary': vendor_summary,
        'delayed_pos': delayed_pos,
        'quantity_errors': quantity_errors,
        'overbilling_cases': overbilling_cases,
        'df': df,
        'spend_trend': spend_trend,
        'vendor_spend': vendor_spend,
        'entity_spend': entity_spend,
        'overbilling_df': overbilling_df,
        'processed': True
    })

# Sidebar navigation options
analysis_option = st.sidebar.selectbox("Select Analysis View", [
    "Total Spend by Service Area",
    "Entity-wise Spend Analysis",
    "Spend by Entity",
    "Total Spend by Material",
    "Top 10 Materials by Spend",
    "Top 10 Vendors by Spend",
    "Spend Distribution by Vendor",
    "Monthly Top Vendors Trend",
    "Top Vendor Monthly Trend",
    "Total PO Order Value & PO Invoice Value Trend",
    "Pending Deliveries by Vendor",
    "Down Payment Analysis by Vendor",
    "Overbilling Analysis",
    "Underbilling Analysis"
])

# Display visualizations based on selected option
if st.session_state.processed:
    if analysis_option == "Total Spend by Service Area":
        st.header("Total Spend by Service Area")
        df_plot = st.session_state.total_spend_by_service_area.sort_values(by='Total PO Ordered Value', ascending=False)
        # Create a pie chart that shows both the INR value and percentage for each service area
        fig = px.pie(df_plot,
                     values='Total PO Ordered Value',
                     names='Service Area',
                     hover_data=['Total PO Invoice Value', 'IT/NON-IT'],
                     hole=0.3)
        fig.update_traces(textinfo='label+percent', 
                          texttemplate="₹%{value:,.0f} (%{percent})")
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Entity-wise Spend Analysis":
        st.header("Entity-wise Spend Analysis")
        entities = st.session_state.df['Entity Name'].unique()
        selected_entity = st.selectbox("Select Entity", sorted(entities))
        df_entity = st.session_state.df[st.session_state.df['Entity Name'] == selected_entity]
        # Group spend by Service Area for the selected entity
        service_area_spend = df_entity.groupby('Service Area').agg({
            'PO Ordered Value in Loc. Curr.': 'sum',
            'PO Invoice Value in Loc. Curr.': 'sum'
        }).reset_index()
        service_area_spend = service_area_spend.sort_values(by='PO Ordered Value in Loc. Curr.', ascending=False)
        
        # Create a grouped bar chart with both Ordered and Invoice values
        fig = go.Figure(data=[
            go.Bar(
                name="PO Ordered Value",
                x=service_area_spend['Service Area'],
                y=service_area_spend['PO Ordered Value in Loc. Curr.'],
                marker_color='indianred',
                text=service_area_spend['PO Ordered Value in Loc. Curr.'],
                texttemplate="₹%{text:,.0f}",
                textposition='auto'
            ),
            go.Bar(
                name="PO Invoice Value",
                x=service_area_spend['Service Area'],
                y=service_area_spend['PO Invoice Value in Loc. Curr.'],
                marker_color='lightsalmon',
                text=service_area_spend['PO Invoice Value in Loc. Curr.'],
                texttemplate="₹%{text:,.0f}",
                textposition='auto'
            )
        ])
        fig.update_layout(
            barmode='group',
            xaxis_title="Service Area",
            yaxis_title="Value (INR)",
            title=f"Spend Analysis for {selected_entity} by Service Area",
            yaxis_tickprefix="₹",
            yaxis_tickformat=","
        )
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Spend by Entity":
        st.header("Spend by Entity")
        fig = px.pie(st.session_state.entity_spend,
                     names='Entity Name',
                     values='PO Ordered Value in Loc. Curr.',
                     title='Spend Distribution Across Entities')
        fig.update_layout(
            showlegend=True,
            legend_title_text='Entities',
            uniformtext_minsize=12,
            uniformtext_mode='hide'
        )
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Total Spend by Material":
        st.header("Total Spend by Material")
        df_plot = st.session_state.total_spend_by_material.sort_values(by='Total PO Ordered Value', ascending=False)
        fig = px.treemap(df_plot,
                         path=['Material Description'],
                         values='Total PO Ordered Value',
                         color='IT/NON-IT',
                         hover_data=['Total PO Invoice Value'])
        fig.update_traces(texttemplate="₹%{value:,.0f}")
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Top 10 Materials by Spend":
        st.header("Top 10 Materials by Spend")
        df_plot = st.session_state.top_10_materials.sort_values(by='Total PO Ordered Value', ascending=False)
        fig = px.bar(df_plot,
                     x='Material Description',
                     y='Total PO Ordered Value',
                     color='IT/NON-IT',
                     hover_data=['Total PO Invoice Value'])
        fig.update_layout(xaxis={'categoryorder':'total descending'})
        fig.update_traces(texttemplate="₹%{y:,.0f}", textposition='outside')
        fig.update_yaxes(tickprefix="₹", tickformat=",")
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Top 10 Vendors by Spend":
        st.header("Top 10 Vendors by Spend")
        df_plot = st.session_state.top_10_vendors.sort_values(by='Total PO Ordered Value', ascending=False)
        fig = px.bar(df_plot,
                     x='Vendor Name',
                     y='Total PO Ordered Value',
                     color='IT/NON-IT',
                     orientation='v',
                     hover_data=['Total PO Invoice Value', 'Vendor Number', 'Entity Name'])
        fig.update_layout(xaxis={'categoryorder':'total descending'})
        fig.update_traces(texttemplate="₹%{y:,.0f}", textposition='outside')
        fig.update_yaxes(tickprefix="₹", tickformat=",")
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Spend Distribution by Vendor":
        st.header("Spend Distribution by Vendor")
        fig = px.treemap(st.session_state.vendor_spend,
                         path=['Vendor Name'],
                         values='PO Ordered Value in Loc. Curr.',
                         hover_data=['PO Invoice Value in Loc. Curr.'],
                         title='Vendor Spend Distribution')
        fig.update_layout(
            coloraxis_colorbar_title="Spend Value (INR)",
            margin=dict(t=50, l=25, r=25, b=25)
        )
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Monthly Top Vendors Trend":
        st.header("Monthly Top Vendors Trend")
        all_months = st.session_state.top_10_vendors_monthly['Month'].unique()
        selected_month = st.selectbox("Select Month", sorted(all_months))
        monthly_data = st.session_state.top_10_vendors_monthly[
            st.session_state.top_10_vendors_monthly['Month'] == selected_month
        ].sort_values(by='Total PO Ordered Value', ascending=False)
        
        fig = px.bar(monthly_data,
                     x='Vendor Name',
                     y='Total PO Ordered Value',
                     color='IT/NON-IT',
                     title=f'Top Vendors for {selected_month}',
                     hover_data=['Total PO Invoice Value', 'Entity Name', 'Vendor Number'])
        fig.update_layout(xaxis_title="Vendor",
                         yaxis_title="Total PO Ordered Value",
                         xaxis={'categoryorder':'total descending'})
        fig.update_traces(texttemplate="₹%{y:,.0f}", textposition='outside')
        fig.update_yaxes(tickprefix="₹", tickformat=",")
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Top Vendor Monthly Trend":
        st.header("Top Vendor Monthly Trend")
        # Get the top vendor for each month
        top_vendor_monthly = st.session_state.top_10_vendors_monthly.groupby('Month').first().reset_index()
        
        fig = px.line(top_vendor_monthly,
                      x='Month',
                      y='Total PO Ordered Value',
                      title='Top Vendor Monthly Trend',
                      labels={'Total PO Ordered Value': 'Total PO Ordered Value (INR)', 'Month': 'Month'},
                      hover_data=['Vendor Name', 'Entity Name', 'IT/NON-IT'])
        fig.update_layout(
            xaxis_title='Month',
            yaxis_title='Total PO Ordered Value (INR)',
            yaxis_tickprefix="₹ ",
            yaxis_tickformat=".2f"
        )
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Total PO Order Value & PO Invoice Value Trend":
        st.header("Total PO Order Value & PO Invoice Value Trend")
        fig = px.line(st.session_state.spend_trend,
                      x='Month',
                      y=['PO Ordered Value in Loc. Curr.', 'PO Invoice Value in Loc. Curr.'],
                      title='Monthly Spend Trend',
                      labels={'value': 'Value (INR)', 'Month': 'Month'})
        fig.update_layout(
            xaxis_title='Month',
            yaxis_title='Value (INR)',
            yaxis_tickprefix="₹ ",
            yaxis_tickformat=".2f"
        )
        st.plotly_chart(fig, use_container_width=True)

    elif analysis_option == "Pending Deliveries by Vendor":
        st.header("Pending Deliveries by Vendor")
        df = st.session_state.df
        df['Pending Deliveries'] = df['Ordered Quantity'] - df['Delivery Quantity']
        pending_deliveries = df.groupby('Vendor Name')['Pending Deliveries'].sum().reset_index()
        fig_pd = px.pie(pending_deliveries,
                        values='Pending Deliveries',
                        names='Vendor Name',
                        title="Pending Deliveries by Vendor")
        st.plotly_chart(fig_pd, use_container_width=True)

    elif analysis_option == "Down Payment Analysis by Vendor":
        st.header("Down Payment Analysis by Vendor")
        df = st.session_state.df
        down_payment_summary = df.groupby('Vendor Name')['PO Down Payment'].sum().reset_index()
        fig_dp = px.pie(down_payment_summary,
                        values='PO Down Payment',
                        names='Vendor Name',
                        title="Down Payment Analysis by Vendor")
        st.plotly_chart(fig_dp, use_container_width=True)

    elif analysis_option == "Overbilling Analysis":
        st.header("Overbilling Analysis (Based on PO Value & PO Invoice Value)")
        
        # Top Vendors by Total Overbilling
        vendor_overbilling = st.session_state.overbilling_df.groupby('Vendor Name', as_index=False)['Overbilling Amount'].sum()
        vendor_overbilling = vendor_overbilling.sort_values(by='Overbilling Amount', ascending=False)
        
        # Pie Chart
        fig_pie = px.pie(vendor_overbilling,
                         names='Vendor Name',
                         values='Overbilling Amount',
                         title="Total Overbilling by Vendor (INR)")
        st.plotly_chart(fig_pie, use_container_width=True)

        # Monthly Trend
        month_positive_overbilling = st.session_state.overbilling_df.groupby('Month', as_index=False)['Overbilling Amount'].sum()
        fig_line = px.line(month_positive_overbilling,
                           x='Month',
                           y='Overbilling Amount',
                           title="Monthly Overbilling Trend (INR)",
                           labels={'Month': 'Month', 'Overbilling Amount': 'Total Overbilling Amount (INR)'})
        fig_line.update_yaxes(tickformat=",", exponentformat="none")
        st.plotly_chart(fig_line, use_container_width=True)

    elif analysis_option == "Underbilling Analysis":
        st.header("Underbilling Analysis (Based on PO Value & PO Invoice Value)")
        underbilling_df = st.session_state.df[st.session_state.df['Overbilling Amount'] < 0].copy()
        underbilling_df['Underbilling Amount'] = -underbilling_df['Overbilling Amount']
        
        # Top Vendors by Total Underbilling
        vendor_underbilling = underbilling_df.groupby('Vendor Name', as_index=False)['Underbilling Amount'].sum()
        vendor_underbilling = vendor_underbilling.sort_values(by='Underbilling Amount', ascending=False)
        
        fig_pie_under = px.pie(vendor_underbilling,
                               names='Vendor Name',
                               values='Underbilling Amount',
                               title="Total Underbilling by Vendor (INR)")
        st.plotly_chart(fig_pie_under, use_container_width=True)
        
        # Monthly Trend
        month_underbilling = underbilling_df.groupby('Month', as_index=False)['Underbilling Amount'].sum()
        fig_line_under = px.line(month_underbilling,
                                 x='Month',
                                 y='Underbilling Amount',
                                 title="Monthly Underbilling Trend (INR)",
                                 labels={'Month': 'Month', 'Underbilling Amount': 'Total Underbilling Amount (INR)'})
        fig_line_under.update_yaxes(tickformat=",", exponentformat="none")
        st.plotly_chart(fig_line_under, use_container_width=True)

    # Excel Report Generation
    st.sidebar.markdown("---")
    if st.sidebar.button("Generate Full Report"):
        output_file = BytesIO()
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Existing sheets
            st.session_state.total_spend_by_vendor.to_excel(writer, sheet_name='Total Spend by Vendor', index=False)
            st.session_state.total_spend_by_material.to_excel(writer, sheet_name='Total Spend by Material', index=False)
            st.session_state.total_spend_by_service_area.to_excel(writer, sheet_name='Total Spend by Service Area', index=False)
            st.session_state.top_10_vendors.to_excel(writer, sheet_name='Top 10 Vendors', index=False)
            st.session_state.top_10_materials.to_excel(writer, sheet_name='Top 10 Materials', index=False)
            st.session_state.top_10_vendors_monthly.to_excel(writer, sheet_name='Top 10 Vendors Monthly', index=False)
            st.session_state.vendor_summary.to_excel(writer, sheet_name='Vendor Analysis', index=False)
            st.session_state.delayed_pos.to_excel(writer, sheet_name='Delayed POs', index=False)
            st.session_state.quantity_errors.to_excel(writer, sheet_name='Quantity Errors', index=False)
            
            # New Overbilling Analysis Sheet with Document Date and Delivery Date
            overbilling_analysis = st.session_state.overbilling_df[[
                'Purchasing Document Number', 'Document Date', 'Delivery Date',
                'Vendor Name', 'Vendor Number', 'Entity Name', 'IT/NON-IT', 
                'PO Ordered Value in Loc. Curr.', 'PO Invoice Value in Loc. Curr.', 
                'Overbilling Amount'
            ]].sort_values(by='Overbilling Amount', ascending=False)
            overbilling_analysis.to_excel(writer, sheet_name='Overbilling Analysis', index=False)
            
            # New Underbilling Analysis Sheet with Document Date and Delivery Date
            underbilling_df = st.session_state.df[st.session_state.df['Overbilling Amount'] < 0].copy()
            underbilling_df['Underbilling Amount'] = -underbilling_df['Overbilling Amount']
            underbilling_analysis = underbilling_df[[
                'Purchasing Document Number', 'Document Date', 'Delivery Date',
                'Vendor Name', 'Vendor Number', 'Entity Name', 'IT/NON-IT', 
                'PO Ordered Value in Loc. Curr.', 'PO Invoice Value in Loc. Curr.', 
                'Underbilling Amount'
            ]].sort_values(by='Underbilling Amount', ascending=False)
            underbilling_analysis.to_excel(writer, sheet_name='Underbilling Analysis', index=False)

        st.sidebar.download_button(
            label="Download Excel Report",
            data=output_file.getvalue(),
            file_name="P2P_Analysis_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload an Excel file to begin analysis")

# Additional styling
st.markdown("""
<style>
    .stPlotlyChart {
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .stSelectbox div[data-baseweb="select"] > div {
        border-color: #4a4a4a;
    }
    .stButton button {
        background-color: #4CAF50;
        color: white;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)
