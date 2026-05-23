import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io

# 1. Page Configuration & Theme Styling
st.set_page_config(page_title="EXtRA Data Analyzer", layout="wide")

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
    
    /* Metric Card Styling */
    .metric-box {
        background-color: #f0f4f8;
        border-left: 5px solid #FFCC00;
        padding: 15px;
        border-radius: 5px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .metric-title { font-size: 14px; color: #555; font-weight: bold; }
    .metric-value { font-size: 24px; color: #003366; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# App Header Component
st.markdown('<div style="text-align: center; margin-bottom: 30px;">'
            '<h1>E<span style="font-size: 0.7em; vertical-align: super;">X</span>tRA Data Analyzer</h1>'
            '<p style="color: #666;">Upload your logistics Excel sheet to get automated insights instantly</p>'
            '</div>', unsafe_allow_html=True)

# 2. File Upload Zone
uploaded_file = st.file_uploader("Upload your Excel File (.xlsx)", type=["xlsx"])

# Helper function to generate standardized charts
def create_bar_chart(df, x_col, y_col, title):
    fig = px.bar(df, x=x_col, y=y_col, title=title,
                 color_discrete_sequence=['#003366']) # Blue Bars
    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        xaxis_title=x_col,
        yaxis_title="Unique Orders Count",
        font=dict(color="#003366")
    )
    fig.update_traces(marker_line_color='#FFCC00', marker_line_width=1.5) # Yellow borders
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
                
                # Dynamic column cleaners
                df_hd.columns = df_hd.columns.str.strip()
                
                # Check for BOOK ID
                if 'BOOK ID' in df_hd.columns:
                    # Parse Dates for the 8th requirement
                    if 'Actual Delivery Date' in df_hd.columns:
                        df_hd['Actual Delivery Date'] = pd.to_datetime(df_hd['Actual Delivery Date'], errors='coerce')
                        df_hd['Month'] = df_hd['Actual Delivery Date'].dt.strftime('%Y-%m ( %B )')

                    columns_to_analyze = [
                        'Group Name', 'Region Name', 'Store Code', 'Area Name', 
                        'Technician', 'Driver', 'Truck No', 'Month'
                    ]
                    
                    cols = st.columns(4)
                    col_idx = 0
                    
                    ppt_data['HD'] = {}
                    
                    for col in columns_to_analyze:
                        if col in df_hd.columns or (col == 'Month' and 'Month' in df_hd.columns):
                            display_name = "Actual Delivery Date (By Month)" if col == 'Month' else col
                            
                            # Critical requirement: Eliminate duplicates on BOOK ID (nunique)
                            summary = df_hd.groupby(col)['BOOK ID'].nunique().reset_index()
                            summary.columns = [display_name, 'Unique Orders']
                            summary = summary.sort_values(by='Unique Orders', ascending=False)
                            
                            # Save for PowerPoint
                            ppt_data['HD'][display_name] = summary
                            
                            # Display Metrics and Charts
                            with cols[col_idx % 4]:
                                st.markdown(f"""
                                <div class="metric-box">
                                    <div class="metric-title">{display_name}</div>
                                    <div class="metric-value">{len(summary)} Unique Categories</div>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            st.plotly_chart(create_bar_chart(summary, display_name, 'Unique Orders', f"Orders by {display_name}"), use_container_width=True)
                            col_idx += 1
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
                    cols_conf = st.columns(4)
                    col_idx = 0
                    
                    ppt_data['Confirmation'] = {}
                    
                    for col in conf_columns:
                        if col in df_conf.columns or (col == 'Month' and 'Month' in df_conf.columns):
                            display_name = "Date (By Month)" if col == 'Month' else col
                            
                            # Eliminate duplicates on BOOK ID
                            summary = df_conf.groupby(col)['BOOK ID'].nunique().reset_index()
                            summary.columns = [display_name, 'Unique Orders']
                            summary = summary.sort_values(by='Unique Orders', ascending=False)
                            
                            ppt_data['Confirmation'][display_name] = summary
                            
                            with cols_conf[col_idx % 4]:
                                st.markdown(f"""
                                <div class="metric-box">
                                    <div class="metric-title">{display_name}</div>
                                    <div class="metric-value">{len(summary)} Unique Categories</div>
                                </div>
                                """, unsafe_allow_html=True)
                                
                            st.plotly_chart(create_bar_chart(summary, display_name, 'Unique Orders', f"Orders by {display_name}"), use_container_width=True)
                            col_idx += 1
                else:
                    st.error("Column 'BOOK ID' not found in Confirmation sheet.")

            # ----------------------------------------
            # TAB 3: RETURN SHEET ANALYSIS
            # ----------------------------------------
            with tab3:
                st.header("Return Sheet Dashboard")
                df_ret = pd.read_excel(xls, sheet_name='Return')
                df_ret.columns = df_ret.columns.str.strip()
                
                # Check for either Date or Return Date to define months
                date_col = 'Date' if 'Date' in df_ret.columns else ([c for c in df_ret.columns if 'date' in c.lower()] + [None])[0]
                
                if 'BOOK ID' in df_ret.columns and date_col:
                    df_ret['Parsed Date'] = pd.to_datetime(df_ret[date_col], errors='coerce')
                    df_ret['Month'] = df_ret['Parsed Date'].dt.strftime('%Y-%m ( %B )')
                    
                    # Eliminate duplicates and group by month
                    summary_ret = df_ret.groupby('Month')['BOOK ID'].nunique().reset_index()
                    summary_ret.columns = ['Month', 'Unique Returns']
                    summary_ret = summary_ret.sort_values(by='Month')
                    
                    ppt_data['Return'] = summary_ret
                    
                    st.markdown(f"""
                    <div class="metric-box" style="max-width: 300px;">
                        <div class="metric-title">Total Unique Returned Orders</div>
                        <div class="metric-value">{summary_ret['Unique Returns'].sum()}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.plotly_chart(create_bar_chart(summary_ret, 'Month', 'Unique Returns', "Returned Orders Trend by Month"), use_container_width=True)
                else:
                    st.error("Make sure 'Return' sheet has both 'BOOK ID' and a valid 'Date' column.")

            # ----------------------------------------
            # POWERPOINT EXPORT LOGIC
            # ----------------------------------------
            st.divider()
            st.subheader("Export Results")
            
            if st.button("Generate & Download PowerPoint Report 📊", use_container_width=True):
                prs = Presentation()
                
                # Color Palette definitions for PPT
                DARK_BLUE = RGBColor(0, 51, 102)
                YELLOW = RGBColor(255, 204, 0)
                
                # Slide 1: Title Slide
                slide_layout = prs.slide_layouts[6] # Blank slide layout
                slide = prs.slides.add_slide(slide_layout)
                
                # Title Text
                tx_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
                tf = tx_box.text_frame
                p = tf.add_paragraph()
                p.text = "EXtRA Data Analysis Report"
                p.font.size = Pt(44)
                p.font.bold = True
                p.font.color.rgb = DARK_BLUE
                
                p2 = tf.add_paragraph()
                p2.text = "Automated Logistics & Performance Summary"
                p2.font.size = Pt(20)
                p2.font.color.rgb = YELLOW
                
                # Dynamic generation of slides based on tables
                for sheet_name, categories in ppt_data.items():
                    if sheet_name in ['HD', 'Confirmation']:
                        for cat_name, df_summary in categories.items():
                            slide = prs.slides.add_slide(prs.slide_layouts[6])
                            
                            # Add Title to Slide
                            tx = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
                            tx.text_frame.text = f"{sheet_name} Sheet - {cat_name} Analysis"
                            tx.text_frame.paragraphs[0].font.size = Pt(24)
                            tx.text_frame.paragraphs[0].font.bold = True
                            tx.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE
                            
                            # Add data table (Top 10 items for visibility)
                            df_top = df_summary.head(10)
                            rows = len(df_top) + 1
                            table_shape = slide.shapes.add_table(rows, 2, Inches(1), Inches(1.5), Inches(8), Inches(4))
                            table = table_shape.table
                            
                            # Header
                            table.cell(0, 0).text = str(df_top.columns[0])
                            table.cell(0, 1).text = str(df_top.columns[1])
                            table.cell(0, 0).fill.solid()
                            table.cell(0, 0).fill.fore_color.rgb = DARK_BLUE
                            table.cell(0, 1).fill.solid()
                            table.cell(0, 1).fill.fore_color.rgb = DARK_BLUE
                            
                            # Populate rows
                            for r_idx, row in df_top.iterrows():
                                table.cell(r_idx + 1, 0).text = str(row.iloc[0])
                                table.cell(r_idx + 1, 1).text = str(row.iloc[1])
                                
                    elif sheet_name == 'Return':
                        slide = prs.slides.add_slide(prs.slide_layouts[6])
                        tx = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
                        tx.text_frame.text = "Return Sheet - Monthly Summary"
                        tx.text_frame.paragraphs[0].font.size = Pt(24)
                        tx.text_frame.paragraphs[0].font.bold = True
                        tx.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE
                        
                        rows = len(categories) + 1
                        table_shape = slide.shapes.add_table(rows, 2, Inches(1), Inches(1.5), Inches(8), Inches(4))
                        table = table_shape.table
                        table.cell(0, 0).text = "Month"
                        table.cell(0, 1).text = "Unique Returns"
                        table.cell(0, 0).fill.solid()
                        table.cell(0, 0).fill.fore_color.rgb = DARK_BLUE
                        table.cell(0, 1).fill.solid()
                        table.cell(0, 1).fill.fore_color.rgb = DARK_BLUE
                        
                        for r_idx, row in categories.iterrows():
                            table.cell(r_idx + 1, 0).text = str(row.iloc[0])
                            table.cell(r_idx + 1, 1).text = str(row.iloc[1])

                # Save Presentation to Memory buffer
                ppt_buffer = io.BytesIO()
                prs.save(ppt_buffer)
                ppt_buffer.seek(0)
                
                # Provide download button
                st.download_button(
                    label="📥 Click here to save the PowerPoint file",
                    data=ppt_buffer,
                    file_name="EXtRA_Data_Analysis.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.success("PowerPoint layout compiled successfully!")
                
    except Exception as e:
        st.error(f"An error occurred while parsing the file: {e}")