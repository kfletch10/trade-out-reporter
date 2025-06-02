import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import base64

# Page configuration
st.set_page_config(
    page_title="Trade Out Report Generator",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f4e79;
        text-align: center;
        margin-bottom: 2rem;
        font-weight: bold;
    }
    .metric-container {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #1f4e79;
        margin: 0.5rem 0;
    }
    .section-header {
        color: #1f4e79;
        font-size: 1.5rem;
        font-weight: bold;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

def load_excel_file(uploaded_file):
    """Load Excel file and return dataframe"""
    try:
        if uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file)
            return df
        else:
            st.error("Please upload an Excel file (.xlsx or .xls)")
            return None
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        return None

def clean_dataframe(df):
    """Clean and prepare dataframe"""
    if df is None:
        return None
    
    # Remove completely empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # Clean column names
    df.columns = df.columns.astype(str).str.strip()
    
    return df

def process_renovation_data(df):
    """Process renovation tracking data"""
    if df is None:
        return None
    
    # Look for key columns (flexible matching)
    unit_col = None
    status_col = None
    year_col = None
    
    for col in df.columns:
        col_lower = str(col).lower()
        if 'unit' in col_lower and unit_col is None:
            unit_col = col
        elif 'status' in col_lower and status_col is None:
            status_col = col
        elif 'year' in col_lower or 'complete' in col_lower and year_col is None:
            year_col = col
    
    if unit_col is None:
        st.error("Could not find 'Unit' column in renovation data")
        return None
    
    # Create processed dataframe
    processed_df = pd.DataFrame()
    processed_df['Unit'] = df[unit_col]
    
    if status_col:
        processed_df['Status'] = df[status_col]
    if year_col:
        processed_df['Year_Complete'] = df[year_col]
    
    # Remove rows where Unit is NaN or not numeric
    processed_df = processed_df[pd.to_numeric(processed_df['Unit'], errors='coerce').notna()]
    processed_df['Unit'] = processed_df['Unit'].astype(int)
    
    return processed_df

def process_rent_roll(df, file_type):
    """Process rent roll data (historical or current)"""
    if df is None:
        return None
    
    # Look for key columns
    unit_col = None
    rent_col = None
    
    for col in df.columns:
        col_lower = str(col).lower()
        if 'unit' in col_lower and unit_col is None:
            unit_col = col
        elif any(word in col_lower for word in ['rent', 'amount', 'price']) and rent_col is None:
            rent_col = col
    
    if unit_col is None or rent_col is None:
        st.error(f"Could not find required columns in {file_type} rent roll")
        return None
    
    # Create processed dataframe
    processed_df = pd.DataFrame()
    processed_df['Unit'] = df[unit_col]
    processed_df[f'{file_type}_Rent'] = pd.to_numeric(df[rent_col], errors='coerce')
    
    # Remove invalid data
    processed_df = processed_df.dropna()
    processed_df['Unit'] = processed_df['Unit'].astype(int)
    
    return processed_df

def calculate_trade_out_metrics(historical_rent, current_rent, renovation_data):
    """Calculate trade out report metrics"""
    
    # Merge all data
    trade_out_df = renovation_data.copy()
    
    # Add rent data
    if historical_rent is not None:
        trade_out_df = trade_out_df.merge(historical_rent, on='Unit', how='left')
    
    if current_rent is not None:
        trade_out_df = trade_out_df.merge(current_rent, on='Unit', how='left')
    
    # Calculate metrics for renovated units only
    renovated_units = trade_out_df[trade_out_df.get('Status', '').str.contains('Done', case=False, na=False)]
    
    if len(renovated_units) == 0:
        return None, "No completed renovations found"
    
    # Calculate rent increases
    if 'Historical_Rent' in renovated_units.columns and 'Current_Rent' in renovated_units.columns:
        renovated_units = renovated_units.dropna(subset=['Historical_Rent', 'Current_Rent'])
        renovated_units['Rent_Increase_Dollar'] = renovated_units['Current_Rent'] - renovated_units['Historical_Rent']
        renovated_units['Rent_Increase_Percent'] = (renovated_units['Rent_Increase_Dollar'] / renovated_units['Historical_Rent']) * 100
        renovated_units['Annual_Income_Increase'] = renovated_units['Rent_Increase_Dollar'] * 12
    
    return renovated_units, None

def create_summary_metrics(df):
    """Create summary metrics for the dashboard"""
    if df is None or len(df) == 0:
        return {}
    
    metrics = {
        'total_renovated_units': len(df),
        'avg_historical_rent': df.get('Historical_Rent', pd.Series()).mean(),
        'avg_current_rent': df.get('Current_Rent', pd.Series()).mean(),
        'avg_rent_increase_dollar': df.get('Rent_Increase_Dollar', pd.Series()).mean(),
        'avg_rent_increase_percent': df.get('Rent_Increase_Percent', pd.Series()).mean(),
        'total_annual_income_increase': df.get('Annual_Income_Increase', pd.Series()).sum(),
        'median_rent_increase_dollar': df.get('Rent_Increase_Dollar', pd.Series()).median(),
        'median_rent_increase_percent': df.get('Rent_Increase_Percent', pd.Series()).median(),
    }
    
    return metrics

def create_visualizations(df, metrics):
    """Create visualizations for the report"""
    charts = {}
    
    if df is None or len(df) == 0:
        return charts
    
    # 1. Rent Increase Distribution
    if 'Rent_Increase_Dollar' in df.columns:
        fig_hist = px.histogram(
            df, 
            x='Rent_Increase_Dollar',
            title='Distribution of Rent Increases ($)',
            labels={'Rent_Increase_Dollar': 'Rent Increase ($)', 'count': 'Number of Units'},
            color_discrete_sequence=['#1f4e79']
        )
        fig_hist.update_layout(showlegend=False)
        charts['rent_increase_distribution'] = fig_hist
    
    # 2. Before vs After Rent Comparison
    if 'Historical_Rent' in df.columns and 'Current_Rent' in df.columns:
        fig_scatter = px.scatter(
            df,
            x='Historical_Rent',
            y='Current_Rent',
            title='Historical vs Current Rent',
            labels={'Historical_Rent': 'Historical Rent ($)', 'Current_Rent': 'Current Rent ($)'},
            color_discrete_sequence=['#1f4e79']
        )
        # Add diagonal line
        max_rent = max(df['Historical_Rent'].max(), df['Current_Rent'].max())
        fig_scatter.add_shape(
            type="line",
            x0=0, y0=0, x1=max_rent, y1=max_rent,
            line=dict(color="red", width=2, dash="dash"),
        )
        charts['before_after_comparison'] = fig_scatter
    
    # 3. Rent Increases by Year
    if 'Year_Complete' in df.columns and 'Rent_Increase_Dollar' in df.columns:
        yearly_data = df.groupby('Year_Complete')['Rent_Increase_Dollar'].agg(['mean', 'count']).reset_index()
        fig_year = px.bar(
            yearly_data,
            x='Year_Complete',
            y='mean',
            title='Average Rent Increase by Completion Year',
            labels={'mean': 'Average Rent Increase ($)', 'Year_Complete': 'Year Completed'},
            color_discrete_sequence=['#1f4e79']
        )
        charts['yearly_performance'] = fig_year
    
    return charts

def generate_pdf_report(df, metrics, property_name="Property"):
    """Generate PDF report"""
    buffer = io.BytesIO()
    
    # Create PDF document
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18)
    
    # Define styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=1,  # Center alignment
        textColor=colors.HexColor('#1f4e79')
    )
    
    # Content list
    content = []
    
    # Title
    content.append(Paragraph(f"{property_name} - Trade Out Report", title_style))
    content.append(Paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y')}", styles['Normal']))
    content.append(Spacer(1, 20))
    
    # Executive Summary
    content.append(Paragraph("Executive Summary", styles['Heading2']))
    
    summary_data = [
        ['Metric', 'Value'],
        ['Total Renovated Units', f"{metrics.get('total_renovated_units', 0):,}"],
        ['Average Historical Rent', f"${metrics.get('avg_historical_rent', 0):,.2f}"],
        ['Average Current Rent', f"${metrics.get('avg_current_rent', 0):,.2f}"],
        ['Average Rent Increase', f"${metrics.get('avg_rent_increase_dollar', 0):,.2f}"],
        ['Average % Increase', f"{metrics.get('avg_rent_increase_percent', 0):.1f}%"],
        ['Total Annual Income Increase', f"${metrics.get('total_annual_income_increase', 0):,.2f}"],
    ]
    
    summary_table = Table(summary_data)
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4e79')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    content.append(summary_table)
    content.append(Spacer(1, 20))
    
    # Unit Details
    if df is not None and len(df) > 0:
        content.append(Paragraph("Unit Details", styles['Heading2']))
        
        # Prepare unit data for table
        unit_columns = ['Unit', 'Historical_Rent', 'Current_Rent', 'Rent_Increase_Dollar', 'Rent_Increase_Percent']
        available_columns = [col for col in unit_columns if col in df.columns]
        
        if available_columns:
            unit_data = [['Unit #', 'Historical Rent', 'Current Rent', 'Increase ($)', 'Increase (%)']]
            
            for _, row in df.head(20).iterrows():  # Limit to first 20 units
                unit_row = [
                    str(int(row['Unit'])),
                    f"${row.get('Historical_Rent', 0):,.2f}",
                    f"${row.get('Current_Rent', 0):,.2f}",
                    f"${row.get('Rent_Increase_Dollar', 0):,.2f}",
                    f"{row.get('Rent_Increase_Percent', 0):.1f}%"
                ]
                unit_data.append(unit_row)
            
            unit_table = Table(unit_data)
            unit_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4e79')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            content.append(unit_table)
    
    # Build PDF
    doc.build(content)
    buffer.seek(0)
    return buffer

# Main Streamlit App
def main():
    st.markdown('<h1 class="main-header">🏢 Trade Out Report Generator</h1>', unsafe_allow_html=True)
    
    # Sidebar for file uploads
    st.sidebar.header("📁 Upload Files")
    
    property_name = st.sidebar.text_input("Property Name", value="Memorial Village")
    
    st.sidebar.markdown("### Required Files")
    
    # File uploads
    historical_rent_file = st.sidebar.file_uploader(
        "Historical Rent Roll", 
        type=['xlsx', 'xls'],
        help="Upload the rent roll from before renovations"
    )
    
    current_rent_file = st.sidebar.file_uploader(
        "Current Rent Roll", 
        type=['xlsx', 'xls'],
        help="Upload the current rent roll"
    )
    
    renovation_file = st.sidebar.file_uploader(
        "Renovation Tracking", 
        type=['xlsx', 'xls'],
        help="Upload the file tracking which units were renovated"
    )
    
    # Initialize session state
    if 'trade_out_data' not in st.session_state:
        st.session_state.trade_out_data = None
    if 'metrics' not in st.session_state:
        st.session_state.metrics = {}
    
    # Process files when uploaded
    if st.sidebar.button("🔄 Process Files", type="primary"):
        if not all([historical_rent_file, current_rent_file, renovation_file]):
            st.error("Please upload all three files before processing.")
            return
        
        with st.spinner("Processing files..."):
            # Load and process files
            historical_df = load_excel_file(historical_rent_file)
            current_df = load_excel_file(current_rent_file)
            renovation_df = load_excel_file(renovation_file)
            
            if all([historical_df is not None, current_df is not None, renovation_df is not None]):
                # Clean and process data
                historical_clean = clean_dataframe(historical_df)
                current_clean = clean_dataframe(current_df)
                renovation_clean = clean_dataframe(renovation_df)
                
                # Process each file type
                historical_processed = process_rent_roll(historical_clean, "Historical")
                current_processed = process_rent_roll(current_clean, "Current")
                renovation_processed = process_renovation_data(renovation_clean)
                
                if all([historical_processed is not None, current_processed is not None, renovation_processed is not None]):
                    # Calculate trade out metrics
                    trade_out_data, error = calculate_trade_out_metrics(
                        historical_processed, current_processed, renovation_processed
                    )
                    
                    if error:
                        st.error(error)
                    else:
                        st.session_state.trade_out_data = trade_out_data
                        st.session_state.metrics = create_summary_metrics(trade_out_data)
                        st.success("✅ Files processed successfully!")
    
    # Display results
    if st.session_state.trade_out_data is not None:
        df = st.session_state.trade_out_data
        metrics = st.session_state.metrics
        
        # Main metrics dashboard
        st.markdown('<div class="section-header">📊 Trade Out Summary</div>', unsafe_allow_html=True)
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "Total Renovated Units",
                f"{metrics.get('total_renovated_units', 0):,}",
                help="Number of completed renovation units"
            )
        
        with col2:
            st.metric(
                "Avg Rent Increase",
                f"${metrics.get('avg_rent_increase_dollar', 0):,.2f}",
                f"{metrics.get('avg_rent_increase_percent', 0):.1f}%",
                help="Average dollar and percentage rent increase"
            )
        
        with col3:
            st.metric(
                "Total Annual Income Increase",
                f"${metrics.get('total_annual_income_increase', 0):,.2f}",
                help="Total additional annual rental income from renovations"
            )
        
        with col4:
            st.metric(
                "Median Rent Increase",
                f"${metrics.get('median_rent_increase_dollar', 0):,.2f}",
                help="Median rent increase across all renovated units"
            )
        
        # Visualizations
        st.markdown('<div class="section-header">📈 Performance Analysis</div>', unsafe_allow_html=True)
        
        charts = create_visualizations(df, metrics)
        
        if charts:
            # Display charts in columns
            if 'rent_increase_distribution' in charts:
                st.plotly_chart(charts['rent_increase_distribution'], use_container_width=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                if 'before_after_comparison' in charts:
                    st.plotly_chart(charts['before_after_comparison'], use_container_width=True)
            
            with col2:
                if 'yearly_performance' in charts:
                    st.plotly_chart(charts['yearly_performance'], use_container_width=True)
        
        # Detailed data table
        st.markdown('<div class="section-header">📋 Unit Details</div>', unsafe_allow_html=True)
        
        # Display options
        col1, col2 = st.columns([3, 1])
        with col2:
            show_all = st.checkbox("Show all units", value=False)
        
        display_df = df if show_all else df.head(20)
        
        # Format the dataframe for display
        display_columns = []
        column_config = {}
        
        if 'Unit' in display_df.columns:
            display_columns.append('Unit')
            column_config['Unit'] = st.column_config.NumberColumn("Unit #", format="%d")
        
        if 'Historical_Rent' in display_df.columns:
            display_columns.append('Historical_Rent')
            column_config['Historical_Rent'] = st.column_config.NumberColumn("Historical Rent", format="$%.2f")
        
        if 'Current_Rent' in display_df.columns:
            display_columns.append('Current_Rent')
            column_config['Current_Rent'] = st.column_config.NumberColumn("Current Rent", format="$%.2f")
        
        if 'Rent_Increase_Dollar' in display_df.columns:
            display_columns.append('Rent_Increase_Dollar')
            column_config['Rent_Increase_Dollar'] = st.column_config.NumberColumn("Rent Increase ($)", format="$%.2f")
        
        if 'Rent_Increase_Percent' in display_df.columns:
            display_columns.append('Rent_Increase_Percent')
            column_config['Rent_Increase_Percent'] = st.column_config.NumberColumn("Rent Increase (%)", format="%.1f%%")
        
        if 'Annual_Income_Increase' in display_df.columns:
            display_columns.append('Annual_Income_Increase')
            column_config['Annual_Income_Increase'] = st.column_config.NumberColumn("Annual Income Increase", format="$%.2f")
        
        if 'Year_Complete' in display_df.columns:
            display_columns.append('Year_Complete')
            column_config['Year_Complete'] = st.column_config.NumberColumn("Year Completed", format="%d")
        
        st.dataframe(
            display_df[display_columns],
            column_config=column_config,
            use_container_width=True,
            hide_index=True
        )
        
        # Export options
        st.markdown('<div class="section-header">📤 Export Report</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Excel export
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                # Summary sheet
                summary_df = pd.DataFrame({
                    'Metric': [
                        'Total Renovated Units',
                        'Average Historical Rent',
                        'Average Current Rent',
                        'Average Rent Increase ($)',
                        'Average Rent Increase (%)',
                        'Median Rent Increase ($)',
                        'Total Annual Income Increase'
                    ],
                    'Value': [
                        metrics.get('total_renovated_units', 0),
                        f"${metrics.get('avg_historical_rent', 0):.2f}",
                        f"${metrics.get('avg_current_rent', 0):.2f}",
                        f"${metrics.get('avg_rent_increase_dollar', 0):.2f}",
                        f"{metrics.get('avg_rent_increase_percent', 0):.1f}%",
                        f"${metrics.get('median_rent_increase_dollar', 0):.2f}",
                        f"${metrics.get('total_annual_income_increase', 0):.2f}"
                    ]
                })
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Unit details sheet
                df.to_excel(writer, sheet_name='Unit Details', index=False)
            
            excel_buffer.seek(0)
            
            st.download_button(
                label="📊 Download Excel Report",
                data=excel_buffer,
                file_name=f"{property_name}_TradeOut_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            # CSV export
            csv_buffer = io.StringIO()
            df.to_csv(csv_buffer, index=False)
            
            st.download_button(
                label="📄 Download CSV Data",
                data=csv_buffer.getvalue(),
                file_name=f"{property_name}_TradeOut_Data_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        
        with col3:
            # PDF export
            if st.button("📋 Generate PDF Report"):
                with st.spinner("Generating PDF report..."):
                    pdf_buffer = generate_pdf_report(df, metrics, property_name)
                    
                    st.download_button(
                        label="📋 Download PDF Report",
                        data=pdf_buffer,
                        file_name=f"{property_name}_TradeOut_Report_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf"
                    )
        
        # Additional insights
        st.markdown('<div class="section-header">💡 Key Insights</div>', unsafe_allow_html=True)
        
        insights_col1, insights_col2 = st.columns(2)
        
        with insights_col1:
            st.markdown("### 🎯 Performance Highlights")
            
            if metrics.get('avg_rent_increase_percent', 0) > 0:
                st.success(f"✅ Average rent increase of {metrics.get('avg_rent_increase_percent', 0):.1f}% achieved")
            
            if metrics.get('total_annual_income_increase', 0) > 0:
                st.success(f"💰 Total additional annual income: ${metrics.get('total_annual_income_increase', 0):,.2f}")
            
            # Top performing units
            if 'Rent_Increase_Percent' in df.columns:
                top_performers = df.nlargest(3, 'Rent_Increase_Percent')
                st.markdown("**Top Performing Units:**")
                for _, unit in top_performers.iterrows():
                    st.write(f"• Unit {int(unit['Unit'])}: {unit['Rent_Increase_Percent']:.1f}% increase")
        
        with insights_col2:
            st.markdown("### 📊 Market Analysis")
            
            if 'Year_Complete' in df.columns and 'Rent_Increase_Dollar' in df.columns:
                yearly_performance = df.groupby('Year_Complete')['Rent_Increase_Dollar'].agg(['mean', 'count'])
                
                st.markdown("**Performance by Year:**")
                for year, data in yearly_performance.iterrows():
                    st.write(f"• {int(year)}: ${data['mean']:.2f} avg increase ({int(data['count'])} units)")
            
            # ROI calculation (if renovation costs were available)
            st.info("💡 **Tip**: Include renovation costs in your data to calculate ROI metrics and payback periods.")
    
    else:
        # Instructions when no data is loaded
        st.markdown("""
        ## 🚀 Get Started
        
        Welcome to the Trade Out Report Generator! This tool helps you analyze rent premiums from your multifamily renovation projects.
        
        ### 📋 How to Use:
        
        1. **Upload three Excel files** using the sidebar:
           - **Historical Rent Roll**: Pre-renovation rent data
           - **Current Rent Roll**: Post-renovation rent data  
           - **Renovation Tracking**: List of renovated units with completion dates
        
        2. **Click "Process Files"** to analyze your data
        
        3. **Review the results** including:
           - Key performance metrics
           - Visual charts and graphs
           - Detailed unit-by-unit analysis
        
        4. **Export professional reports** in Excel, CSV, or PDF format
        
        ### 📊 Key Metrics Generated:
        - Average rent increase ($ and %)
        - Total additional annual income
        - Performance by renovation year
        - Unit-by-unit trade out analysis
        - Market insights and trends
        
        ### 📁 File Requirements:
        - Files must be in Excel format (.xlsx or .xls)
        - Must contain columns for Unit numbers and Rent amounts
        - Renovation file should include completion status and dates
        
        Ready to analyze your renovation performance? Upload your files to get started! 🏢
        """)

if __name__ == "__main__":
    main()
