import streamlit as st
import pandas as pd
from datetime import datetime
import utils

st.set_page_config(
    page_title="Job Status Analyzer",
    page_icon="📊",
    layout="wide"
)

# Add custom CSS
st.markdown("""
    <style>
    .stProgress > div > div > div > div {
        background-color: #F63366;
    }
    .stDownloadButton button {
        background-color: #F63366;
        color: white;
        border-radius: 5px;
        padding: 0.5rem 1rem;
    }
    div[data-testid="stDataFrame"] div[role="cell"] {
        font-family: monospace;
    }
    div[data-testid="stDataFrame"] div[role="columnheader"] {
        background-color: #1F4E78;
        color: white;
        font-weight: bold;
    }
    div[data-testid="stDataFrame"] {
        border: 1px solid #E0E0E0;
        border-radius: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# Title and description
st.title("📊 Job Status Analyzer")
st.markdown("""
    Upload CSV files containing job status information to analyze:
    - Total job counts per vessel
    - New job counts
    - Generate formatted Excel reports
""")

# File uploader
uploaded_files = st.file_uploader(
    "Upload CSV files",
    type=['csv'],
    accept_multiple_files=True,
    help="Select one or more CSV files containing job status information"
)

if uploaded_files:
    # Process files
    progress_bar = st.progress(0)
    status_text = st.empty()

    summary_data = []
    for i, file in enumerate(uploaded_files):
        status_text.text(f"Processing {file.name}...")
        summary_data.append(utils.process_csv_file(file))
        progress_bar.progress((i + 1) / len(uploaded_files))

    # Create DataFrame
    df = pd.DataFrame(summary_data)

    # Convert date strings to datetime for filtering
    df['Date Extracted from File Name'] = pd.to_datetime(
        df['Date Extracted from File Name'],
        format='%d-%m-%Y',
        errors='coerce'
    )

    # Filters
    st.subheader("📌 Filters")
    col1, col2 = st.columns(2)

    with col1:
        vessel_filter = st.multiselect(
            "Filter by Vessel Name",
            options=sorted(df['Vessel Name'].unique()),
            help="Select one or more vessels to filter the data"
        )

    with col2:
        min_date = df['Date Extracted from File Name'].min()
        max_date = df['Date Extracted from File Name'].max()
        date_range = st.date_input(
            "Select Date Range",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date()
        )

    # Apply filters
    filtered_df = df.copy()
    if vessel_filter:
        filtered_df = filtered_df[filtered_df['Vessel Name'].isin(vessel_filter)]
    if len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (filtered_df['Date Extracted from File Name'].dt.date >= start_date) &
            (filtered_df['Date Extracted from File Name'].dt.date <= end_date)
        ]

    # Summary Statistics
    st.subheader("📈 Summary Statistics")

    # Overall metrics in a smaller space
    total_metrics = st.columns(2)
    total_metrics[0].metric("Total Files", len(filtered_df))
    total_metrics[1].metric("Total Vessels", filtered_df['Vessel Name'].nunique())

    # Data Visualizations
    st.subheader("📊 Data Visualizations")

    # Create tabs for different visualizations
    tab1, tab2, tab3 = st.tabs([
        "📊 Job Distribution", 
        "📈 Timeline Analysis", 
        "🥧 Jobs Breakdown"
    ])

    with tab1:
        st.plotly_chart(
            utils.create_vessel_job_distribution_chart(filtered_df),
            use_container_width=True
        )

    with tab2:
        st.plotly_chart(
            utils.create_jobs_timeline_chart(filtered_df),
            use_container_width=True
        )

    with tab3:
        st.plotly_chart(
            utils.create_jobs_pie_chart(filtered_df),
            use_container_width=True
        )

    # Per-vessel detailed breakdown with expanders
    st.subheader("📊 Per Vessel File Breakdown")

    # Format the date column to show only the date
    filtered_df_display = filtered_df.copy()
    filtered_df_display['Date Extracted from File Name'] = filtered_df_display['Date Extracted from File Name'].dt.strftime('%d-%m-%Y')

    # Group by vessel
    for vessel in sorted(filtered_df['Vessel Name'].unique()):
        vessel_data = filtered_df_display[filtered_df_display['Vessel Name'] == vessel]

        # Create expander for each vessel
        with st.expander(f"🚢 {vessel} - {len(vessel_data)} files"):
            # Vessel total metrics
            st.markdown(f"**Total Jobs: {vessel_data['Total Count of Jobs'].sum()}** | "
                       f"**New Jobs: {vessel_data['New Job Count'].sum()}**")

            # Individual file details
            st.dataframe(
                vessel_data[['File Name', 'Date Extracted from File Name', 
                           'Total Count of Jobs', 'New Job Count']]
                .sort_values('Date Extracted from File Name', ascending=False),
                use_container_width=True,
                hide_index=True
            )

    # Display full detailed results
    st.subheader("📋 Detailed Results")
    st.dataframe(
        filtered_df_display,
        use_container_width=True,
        hide_index=True
    )

    # Download button for Excel report
    if st.button("📥 Generate Excel Report"):
        excel_file = utils.create_excel_report(filtered_df_display)
        st.download_button(
            label="Download Excel Report",
            data=excel_file,
            file_name=f"Job_Status_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("👆 Please upload your CSV files to begin analysis")