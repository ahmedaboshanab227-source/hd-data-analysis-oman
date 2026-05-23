import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io

# 1. Page Configuration & Theme Styling
st.set_page_config(page_title="EXTRA HD DATA ANALYZER", layout="wide")

# Custom CSS for Official eXtra Theme (Blue & Yellow) with Enhanced Background Watermark
st.markdown("""
    <style>
    /* eXtra Theme Background with Watermark */
    .stApp {
        background-color: #005EA6; /* eXtra Brand Blue */
        background-image: linear-gradient(rgba(0, 94, 166, 0.93), rgba(0, 94, 166, 0.93)), 
                          url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="300" height="300" viewBox="0 0 300 300"><text x="40%" y="50%" font-family="Arial, sans-serif" font-weight="900" font-size="50" fill="rgba(255, 255, 255, 0.03)" text-anchor="middle">extra</text><text x="75%" y="52%" font-family="Arial, sans-serif" font-weight="900" font-size="65" fill="rgba(255, 204, 0, 0.04)" text-anchor="middle">X</text></svg>');
        background-repeat: repeat;
        background-position: center;
    }
    
    /* Typography & Global Color Overrides for Dark Blue Background */
    h1, h2, h3, h4, p, span, label, .stTabs [data-baseweb="tab"] { 
        color: #FFFFFF !important; 
        font-family: 'Arial', sans-serif; 
    }
    
    /* Metric/Tabs Customization */
    .stTabs [data-baseweb="tab"]:aria-selected="true" {
        color: #FFCC00 !important;
        border-bottom-color: #FFCC00 !important;
    }
    
    /* Professional Glassmorphism Container for Data Tables */
    .stDataFrame, div[data-testid="stDataFrame"] {
        background: rgba(255, 255, 255, 0.05) !important;
        border: 2px solid #FFCC00 !important;
        border-radius: 8px !important;
        padding: 5px;
    }
    
    /* File Uploader Customization */
    section[data-testid="stFileUploadDropzone"] {
        background-color: rgba(255, 255, 255, 0.1) !important;
        border: 2px dashed #FFCC00 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# App Header Component - Styled exactly like eXtra Logo concept
st.markdown('<div style="text-align: center; margin-bottom: 30px; padding-top: 20px;">'
            '<h1 style="font-size: 3.2em; font-weight: 900; letter-spacing: 2px; margin-bottom:0px;">إكسترا <span style="color: #FFCC00;">extra</span></h1>'
            '<h3 style="font-size: 1.5em; color: #E0E0E0 !important; margin-top:5px;">HD DATA ANALYZER</h3>'
            '<p style="color: #FFCC00; font-size: 1.1em; font-weight: bold;">Advanced Logistics Dashboard & Automated Insights</p>'
            '</div>', unsafe_allow_html=True)

# 2. File Upload Zone
uploaded_file = st.file_uploader("Upload your Logistics Excel File (.xlsx)", type=["xlsx"])

# Helper function to generate 3D-effect (Isomorphic/Embossed) Bar Charts
def create_3d_bar_chart(df, x_col, y_col, title):
    # Creating a trace with specific rendering properties to simulate 3D depth
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=df[x_col],
        y=df[y_col],
        text=df[y_col],
        textposition='outside',
        textfont=dict(color='#FFFFFF', size=12, family='Arial Black'),
        marker=dict(
            color='#FFCC00', # eXtra Yellow Bars for high contrast
            line=dict(color='#FFFFFF', width=1.5),
            # Pattern and shadow elements mimic 3D extrusion/depth
            pattern=dict(shape="/", solidity=0.1), 
        ),
        # Applying a subtle shadow/3D ridge effect via error bars tweak or layout boundaries
    ))
    
    fig.update_layout(
        title=dict(text=title, font=dict(color='#FFFFFF', size=16, family='Arial')),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(
            title=x_col, 
            titlefont=dict(color='#E0E0E0'), 
            tickfont=dict(color='#FFFFFF'),
            gridcolor='rgba(255,255,255,0.1)'
        ),
        yaxis=dict(
            title="Unique Orders Count", 
            titlefont=dict(color='#E0E0E0'), 
            tickfont=dict(color='#FFFFFF'),
            gridcolor='rgba(255,255,255,0.15)',
            zerolinecolor='rgba(255,255,255,0.3)'
        ),
        margin=dict(l=40, r=40, t=60, b=40),
        bargap=0.3, # Gives a more robust, chunky 3D look to pillars
    )
    return fig

# 3. Processing Core
if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheets_in_file = xls.sheet_names
        
        required_sheets = ['HD', 'Confirmation', 'Return']
        missing_sheets = [s for s in required_sheets if s not in sheets_in_file]
        
        if missing_sheets:
            st.error(f"Error: Missing required sheets: {', '.join(missing_sheets)}")
        else:
            ppt_data = {}

            # Navigation Tabs
            tab1, tab2, tab3 = st.tabs(["📊 HD Analysis", "✅ Confirmation Analysis", "🔄 Return Analysis"])

            # ----------------------------------------
            # TAB 1: HD SHEET ANALYSIS
            # ----------------------------------------
            with tab1:
                st.markdown("<h2 style='color:#FFCC00 !important;'>HD Sheet Dashboard</h2>", unsafe_allow_html=True)
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
                            
                            summary = df_hd.groupby(col)['BOOK ID'].nunique().reset_index()
                            summary.columns = [display_name, 'Unique Orders']
                            summary = summary.sort_values(by='Unique Orders', ascending=False)
                            
                            ppt_data['HD'][display_name] = summary
                            
                            # Layout Split
                            chart_col, table_col = st.columns([2, 1])
                            with chart_col:
                                st.plotly_chart(create_3d_bar_chart(summary, display_name, 'Unique Orders', f"Orders by {display_name}"), use_container_width=True)
                            with table_col:
                                st.markdown(f"<p style='color:#FFCC00; font-weight:bold;'>📌 {display_name} Summary:</p>", unsafe_allow_html=True)
                                st.dataframe(summary, hide_index=True, use_container_width=True)
                                
                            st.markdown("<hr style='border-color: rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
                else:
                    st.error("Column 'BOOK ID' not found in HD sheet.")

            # ----------------------------------------
            # TAB 2: CONFIRMATION SHEET ANALYSIS
            # ----------------------------------------
            with tab2:
                st.markdown("<h2 style='color:#FFCC00 !important;'>Confirmation Sheet Dashboard</h2>", unsafe_allow_html=True)
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
                            
                            chart_col, table_col = st.columns([2, 1])
                            with chart_col:
                                st.plotly_chart(create_3d_bar_chart(summary, display_name, 'Unique Orders', f"Orders by {display_name}"), use_container_width=True)
                            with table_col:
                                st.markdown(f"<p style='color:#FFCC00; font-weight:bold;'>📌 {display_name} Summary:</p>", unsafe_allow_html=True)
                                st.dataframe(summary, hide_index=True, use_container_width=True)
                                
                            st.markdown("<hr style='border-color: rgba(255,255,255,0.1);'>", unsafe_allow_html=True)
                else:
                    st.error("Column 'BOOK ID' not found in Confirmation sheet.")

            # ----------------------------------------
            # TAB 3: RETURN SHEET ANALYSIS
            # ----------------------------------------
            with tab3:
                st.markdown("<h2 style='color:#FFCC00 !important;'>Return Sheet Dashboard</h2>", unsafe_allow_html=True)
                df_ret = pd.read_excel(xls, sheet_name='Return')
                df_ret.columns = df_ret.columns.str.strip()
                
                date_col = 'Date' if 'Date' in df_ret.columns else ([c for c in df_ret.columns if 'date' in c.lower()] + [None])[0]
                
                if 'BOOK ID' in df_ret.columns and date_col:
                    df_ret['Parsed Date'] = pd.to_datetime(df_ret[date_col], errors='coerce')
                    df_ret['Month'] = df_ret['Parsed Date'].dt.strftime('%Y-%m ( %B )')
                    
                    summary_ret = df_ret.groupby('Month')['BOOK ID'].nunique().reset_index()
                    summary_ret.columns = ['Month', 'Unique Returns']
                    summary_ret = summary_ret.sort_values(by='Month')
                    
                    ppt_data['Return'] = summary_ret
                    
                    chart_col, table_col = st.columns([2, 1])
                    with chart_col:
                        st.plotly_chart(create_3d_bar_chart(summary_ret, 'Month', 'Unique Returns', "Returned Orders Trend by Month"), use_container_width=True)
                    with table_col:
                        st.markdown("<p style='color:#FFCC00; font-weight:bold;'>📌 Monthly Returns Summary:</p>", unsafe_allow_html=True)
                        st.dataframe(summary_ret, hide_index=True, use_container_width=True)
                else:
                    st.error("Make sure 'Return' sheet has both 'BOOK ID' and a valid 'Date' column.")

            # ----------------------------------------
            # POWERPOINT EXPORT LOGIC (Corporate Palette)
            # ----------------------------------------
            st.markdown("<br><br>", unsafe_allow_html=True)
            st.subheader("📥 Export Executive Report")
            
            if st.button("Generate & Download PowerPoint Report 📊", use_container_width=True):
                prs = Presentation()
                DARK_BLUE = RGBColor(0, 94, 166)
                YELLOW = RGBColor(255, 204, 0)
                
                # Title Slide
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                tx_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
                tf = tx_box.text_frame
                p = tf.add_paragraph()
                p.text = "EXTRA HD DATA REPORT"
                p.font.size = Pt(40)
                p.font.bold = True
                p.font.color.rgb = DARK_BLUE
                
                p2 = tf.add_paragraph()
                p2.text = "Automated Logistics Analytics & Performance Summary"
                p2.font.size = Pt(18)
                p2.font.color.rgb = YELLOW
                
                for sheet_name, categories in ppt_data.items():
                    if sheet_name in ['HD', 'Confirmation']:
                        for cat_name, df_summary in categories.items():
                            slide = prs.slides.add_slide(prs.slide_layouts[6])
                            tx = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
                            tx.text_frame.text = f"{sheet_name} Sheet - {cat_name} Analysis"
                            tx.text_frame.paragraphs[0].font.size = Pt(24)
                            tx.text_frame.paragraphs[0].font.bold = True
                            tx.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE
                            
                            df_top = df_summary.head(10)
                            rows = len(df_top) + 1
                            table_shape = slide.shapes.add_table(rows, 2, Inches(1), Inches(1.5), Inches(8), Inches(4))
                            table = table_shape.table
                            
                            table.cell(0, 0).text = str(df_top.columns[0])
                            table.cell(0, 1).text = str(df_top.columns[1])
                            table.cell(0, 0).fill.solid()
                            table.cell(0, 0).fill.fore_color.rgb = DARK_BLUE
                            table.cell(0, 1).fill.solid()
                            table.cell(0, 1).fill.fore_color.rgb = DARK_BLUE
                            
                            for r_idx, row in df_top.reset_index(drop=True).iterrows():
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
                        
                        for r_idx, row in categories.reset_index(drop=True).iterrows():
                            table.cell(r_idx + 1, 0).text = str(row.iloc[0])
                            table.cell(r_idx + 1, 1).text = str(row.iloc[1])

                ppt_buffer = io.BytesIO()
                prs.save(ppt_buffer)
                ppt_buffer.seek(0)
                
                st.download_button(
                    label="📥 Click here to save the PowerPoint file",
                    data=ppt_buffer,
                    file_name="eXtra_HD_Data_Analysis.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.success("PowerPoint layout compiled successfully!")
                
    except Exception as e:
        st.error(f"An error occurred while parsing the file: {e}")
