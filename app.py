import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io

# 1. Page Configuration & Theme Styling
st.set_page_config(page_title="EXTRA HD DATA ANALYZER", layout="wide")

# Custom CSS for Blue/Yellow theme and the requested watermark background
st.markdown("""
    <style>
    /* Background Watermark */
    .stApp {
        background-image: linear-gradient(rgba(255,255,255,0.92), rgba(255,255,255,0.92)), url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600" viewBox="0 0 800 600"><text x="50%" y="50%" font-family="Arial, sans-serif" font-weight="bold" font-size="140" fill="rgba(0, 51, 102, 0.04)" text-anchor="middle">E<tspan font-size="90" dy="-15">X</tspan>tRA</text></svg>');
        background-repeat: repeat;
        background-position: center;
    }
    
    /* Global Styles */
    h1, h2, h3 { color: #003366 !important; font-family: 'Arial', sans-serif; }
    
    /* Global Table Styling */
    .stDataFrame {
        border: 1px solid #003366;
        border-radius: 5px;
    }
    </style>
    """, unsafe_allow_html=True)

# App Header Component - New Header Style with centered Big Yellow X below title
st.markdown('<div style="text-align: center; margin-bottom: 30px;">'
            '<h1 style="font-size: 3em; letter-spacing: 1px;">EXTRA HD DATA ANALYZER</h1>'
            '<div style="font-size: 5.5em; font-weight: 900; color: #FFCC00; margin-top: -10px; margin-bottom: 5px; line-height: 1em;">X</div>'
            '<p style="color: #666; font-size: 1.1em;">Upload your logistics Excel sheet to get automated insights instantly</p>'
            '</div>', unsafe_allow_html=True)

# 2. File Upload Zone
uploaded_file = st.file_uploader("Upload your Excel File (.xlsx)", type=["xlsx"])

# Helper function to generate standardized charts with text values on top
def create_bar_chart(df, x_col, y_col, title):
    fig = px.bar(df, x=x_col, y=y_col, title=title, text_auto=True,
                 color_discrete_sequence=['#003366']) # Blue Bars
    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        xaxis_title=x_col,
        yaxis_title="Unique Orders Count",
        font=dict(color="#003366")
    )
    fig.update_traces(
        marker_line_color='#FFCC00', 
        marker_line_width=1.5,
        textposition='outside' # Puts numbers clearly on top of bars
    )
    return fig

# 3. Processing Core
if uploaded_file is not None:
    try:
        # Load all requested sheets
        xls = pd.ExcelFile(uploaded_file)
        sheets_in_file = xls.sheet_names
        
        # Check for required sheets
        required_sheets = ['HD', 'Confirmation', 'Return']
        missing_sheets = [s for s in required_sheets if s not in sheets_in_file]
        
        if missing_sheets:
            st.error(f"Error: Missing required sheets: {', '.join(missing_sheets)}")
        else:
            # Container for all processed summary data to export to PowerPoint later
            ppt_data = {}

            # Create Tabs for Navigation
            tab1, tab2, tab3 = st.tabs(["HD Analysis", "Confirmation Analysis", "Return Analysis"])

            # ----------------------------------------
            # TAB 1: HD SHEET ANALYSIS
            # ----------------------------------------
            with tab1:
                st.header("HD Sheet Dashboard")
                df_hd = pd.read_excel(xls, sheet_name='HD')
                df_hd.columns = df_hd.columns.str.strip()
                
                if 'BOOK ID' in df_hd.columns:
                    if 'Actual Delivery Date' in df_hd.columns:
                        df_hd['Actual Delivery Date'] = pd.to_datetime(df_hd['Actual Delivery Date'], errors='coerce')
                        df_hd['Month'] = df_hd['Actual Delivery Date'].dt.strftime('%Y-%m ( %B )')

                    columns_to_analyze = [
                        'Group Name', 'Region Name', 'Store Code', 'Area Name', 
                        'Technician', 'Driver', 'Truck No', 'Month'
                    ]
                    
                    ppt_data['HD'] = {}
                    
                    for col in columns_to_analyze:
                        if col in df_hd.columns or (col == 'Month' and 'Month' in df_hd.columns):
                            display_name = "Actual Delivery Date (By Month)" if col == 'Month' else col
                            
                            # Eliminate duplicates on BOOK ID (nunique)
                            summary = df_hd.groupby(col)['BOOK ID'].nunique().reset_index()
                            summary.columns = [display_name, 'Unique Orders']
                            summary = summary.sort_values(by='Unique Orders', ascending=False)
                            
                            ppt_data['HD'][display_name] = summary
                            
                            # Layout split: Chart on left, Data Table with numbers on right
                            chart_col, table_col = st.columns([2, 1])
                            with chart_col:
                                st.plotly_chart(create_bar_chart(summary, display_name, 'Unique Orders', f"Orders by {display_name}"), use_container_width=True)
                            with table_col:
                                st.markdown(f"**{display_name} Summary Table:**")
                                st.dataframe(summary, hide_index=True, use_container_width=True)
                                
                            st.divider()
                else:
                    st.error("Column 'BOOK ID' not found in HD sheet.")

            # ----------------------------------------
            # TAB 2: CONFIRMATION SHEET ANALYSIS
            # ----------------------------------------
            with tab2:
                st.header("Confirmation Sheet Dashboard")
                df_conf = pd.read_excel(xls, sheet_name='Confirmation')
                df_conf.columns = df_conf.columns.str.strip()
                
                if 'BOOK ID' in df_conf.columns:
                    if 'Date' in df_conf.columns:
                        df_conf['Date'] = pd.to_datetime(df_conf['Date'], errors='coerce')
                        df_conf['Month'] = df_conf['Date'].dt.strftime('%Y-%m ( %B )')
                        
                    conf_columns = ['Technician', 'Driver', 'Truck No', 'Month']
                    ppt_data['Confirmation'] = {}
                    
                    for col in conf_columns:
                        if col in df_conf.columns or (col == 'Month' and 'Month' in df_conf.columns):
                            display_name = "Date (By Month)" if col == 'Month' else col
                            
                            summary = df_conf.groupby(col)['BOOK ID'].nunique().reset_index()
                            summary.columns = [display_name, 'Unique Orders']
                            summary = summary.sort_values(by='Unique Orders', ascending=False)
                            
                            ppt_data['Confirmation'][display_name] = summary
                            
                            # Layout split for numbers table
                            chart_col, table_col = st.columns([2, 1])
                            with chart_col:
                                st.plotly_chart(create_bar_chart(summary, display_name, 'Unique Orders', f"Orders by {display_name}"), use_container_width=True)
                            with table_col:
                                st.markdown(f"**{display_name} Summary Table:**")
                                st.
