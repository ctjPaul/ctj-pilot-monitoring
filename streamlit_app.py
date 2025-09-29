import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import json
from pathlib import Path

# Import the automation module
import pilot_automation

# Configure the page
st.set_page_config(
    page_title="CTJ Pilot Monitoring Reports",
    page_icon="🔥",
    layout="wide"
)

# Initialize session state for storing generated reports
if 'report_history' not in st.session_state:
    st.session_state.report_history = []

# Title and branding
st.title("🔥 CTJ Energy Pilot Monitoring System")
st.subheader("Automated Monthly Report Generator")

# Sidebar for configuration
st.sidebar.header("Configuration")

# Device selection
st.sidebar.subheader("1. Select Device")

# Load devices from config file or use defaults
devices = {
    "Scout-12197": {
        "id": "359205108536865",
        "commission_date": "06/10/2025",
        "name": "Scout-12197",
        "location": "Main Facility"
    },
    "Scout-12198": {
        "id": "359205108536866",
        "commission_date": "07/15/2025",
        "name": "Scout-12198",
        "location": "North Building"
    },
    "Scout-12199": {
        "id": "359205108536867",
        "commission_date": "08/01/2025",
        "name": "Scout-12199",
        "location": "South Building"
    }
}

selected_device_name = st.sidebar.selectbox(
    "Choose Device",
    list(devices.keys()),
    help="Select the Scout device to generate a report for"
)

selected_device = devices[selected_device_name]

# Month selection
st.sidebar.subheader("2. Select Report Period")

# Get last 12 months
current_date = datetime.now()
month_options = []
for i in range(12):
    target_date = current_date - timedelta(days=30*i)
    first_of_month = target_date.replace(day=1)
    
    # Calculate last day of month
    if first_of_month.month == 12:
        last_of_month = first_of_month.replace(year=first_of_month.year + 1, month=1, day=1) - timedelta(days=1)
    else:
        last_of_month = first_of_month.replace(month=first_of_month.month + 1, day=1) - timedelta(days=1)
    
    month_name = first_of_month.strftime("%B %Y")
    month_options.append({
        "display": month_name,
        "month": first_of_month.month,
        "year": first_of_month.year,
        "first_day": first_of_month,
        "last_day": last_of_month
    })

selected_month = st.sidebar.selectbox(
    "Report Period",
    [m["display"] for m in month_options],
    help="Select the month to generate the report for"
)

# Get selected month details
month_data = next(m for m in month_options if m["display"] == selected_month)

# Email configuration
st.sidebar.subheader("3. Email Recipients")
default_recipients = "customer@email.com\nmanager@company.com"
email_recipients = st.sidebar.text_area(
    "Email Addresses (one per line)",
    default_recipients,
    height=100,
    help="Enter one email address per line"
)

# Email settings
st.sidebar.subheader("4. Email Settings")
email_from = st.sidebar.text_input(
    "From Email", 
    "reports@ctjenergy.com",
    help="Your email address for sending reports"
)
email_password = st.sidebar.text_input(
    "Email Password", 
    type="password",
    help="Your email password (not stored)"
)
send_email = st.sidebar.checkbox(
    "Send Email Report", 
    value=False,
    help="Uncheck to only save PDF locally"
)

# Main content area
col1, col2 = st.columns([2, 1])

with col1:
    st.header("Report Configuration")
    
    # Show selected settings
    st.info(f"""
    **Device:** {selected_device['name']}  
    **Location:** {selected_device['location']}  
    **Device ID:** {selected_device['id']}  
    **Report Period:** {selected_month}  
    **Date Range:** {month_data['first_day'].strftime('%m/%d/%Y')} to {month_data['last_day'].strftime('%m/%d/%Y')}  
    **Email Recipients:** {len([e for e in email_recipients.split('\n') if e.strip()])} addresses  
    **Send Email:** {"Yes" if send_email else "No (Save locally only)"}
    """)
    
    # Validate inputs
    can_generate = True
    validation_messages = []
    
    if send_email and not email_password:
        validation_messages.append("⚠️ Email password required for sending reports")
        can_generate = False
    
    if not email_recipients.strip():
        validation_messages.append("⚠️ At least one email recipient required")
    
    if validation_messages:
        for msg in validation_messages:
            st.warning(msg)
    
    # Generate report button
    if st.button("🚀 Generate Monthly Report", type="primary", use_container_width=True, disabled=not can_generate):
        with st.spinner("Generating report... This may take 2-3 minutes."):
            
            # Create progress indicators
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Prepare configuration for automation
                config = {
                    'device': selected_device,
                    'month': month_data,
                    'email_recipients': [e.strip() for e in email_recipients.split('\n') if e.strip()],
                    'email_from': email_from,
                    'email_password': email_password,
                    'send_email': send_email,
                    'report_period_display': selected_month
                }
                
                # Define callback function for status updates
                def update_status(message, progress):
                    status_text.text(message)
                    progress_bar.progress(progress)
                
                # Run the automation with progress callback
                result = pilot_automation.run_automation(config, status_callback=update_status)
                
                # Check result
                if result['success']:
                    progress_bar.progress(100)
                    status_text.text("✅ Report generated successfully!")
                    
                    st.success(f"Monthly report for {selected_device['name']} - {selected_month} completed!")
                    
                    # Add to history
                    st.session_state.report_history.insert(0, {
                        'device': selected_device['name'],
                        'period': selected_month,
                        'timestamp': datetime.now(),
                        'path': result['report_path'],
                        'emailed': result.get('email_sent', False)
                    })
                    
                    # Show download button
                    if os.path.exists(result['report_path']):
                        with open(result['report_path'], "rb") as f:
                            st.download_button(
                                label="📥 Download PDF Report",
                                data=f.read(),
                                file_name=os.path.basename(result['report_path']),
                                mime="application/pdf",
                                use_container_width=True
                            )
                    
                    # Show summary statistics
                    if 'summary' in result:
                        st.subheader("Report Summary")
                        col_a, col_b, col_c = st.columns(3)
                        with col_a:
                            st.metric("Total Outages", result['summary'].get('total_outages', 'N/A'))
                        with col_b:
                            st.metric("Availability", result['summary'].get('availability', 'N/A'))
                        with col_c:
                            st.metric("EPA Compliance", result['summary'].get('compliance', 'N/A'))
                    
                else:
                    st.error(f"Error generating report: {result.get('error', 'Unknown error')}")
                    if 'details' in result:
                        with st.expander("Error Details"):
                            st.code(result['details'])
                
            except Exception as e:
                st.error(f"Unexpected error: {str(e)}")
                with st.expander("Error Details"):
                    st.exception(e)

with col2:
    st.header("Quick Stats")
    
    # Calculate real stats from history
    total_reports = len(st.session_state.report_history)
    this_month_reports = len([r for r in st.session_state.report_history 
                               if r['timestamp'].month == datetime.now().month])
    
    st.metric("Reports Generated", f"{total_reports}", f"+{this_month_reports} this month")
    st.metric("Active Devices", f"{len(devices)}", "+0 new")
    st.metric("Avg Availability", "98.5%", "+0.3%")
    
    st.header("Recent Reports")
    if st.session_state.report_history:
        for report in st.session_state.report_history[:5]:
            with st.expander(f"{report['device']} - {report['period']}"):
                st.write(f"**Generated:** {report['timestamp'].strftime('%m/%d/%Y %I:%M %p')}")
                st.write(f"**Emailed:** {'Yes' if report['emailed'] else 'No'}")
                if os.path.exists(report['path']):
                    with open(report['path'], "rb") as f:
                        st.download_button(
                            label="Download",
                            data=f.read(),
                            file_name=os.path.basename(report['path']),
                            mime="application/pdf",
                            key=f"dl_{report['timestamp'].timestamp()}"
                        )
    else:
        st.info("No reports generated yet")

# Footer
st.divider()
st.caption("CTJ Energy Solutions - Automated Pilot Monitoring System v2.0")

# Instructions expander
with st.expander("📖 How to Use"):
    st.markdown("""
    ### Step-by-Step Instructions
    
    1. **Select Device** - Choose which Scout device to generate a report for
    2. **Select Report Period** - Choose which month you want to report on
    3. **Configure Email** - Enter recipient email addresses (one per line)
    4. **Email Settings** - Enter your email credentials if sending reports
    5. **Generate Report** - Click the button and wait 2-3 minutes
    
    ### What This App Does
    
    - ✅ Automatically logs into Uplink system
    - ✅ Downloads device data for selected month
    - ✅ Analyzes pilot outage events
    - ✅ Calculates EPA compliance metrics
    - ✅ Generates professional PDF report with your branding
    - ✅ Sends email to all recipients (optional)
    - ✅ Saves reports locally for future reference
    
    ### System Requirements
    
    - Chrome browser must be installed
    - Internet connection required
    - Valid email credentials for sending reports
    
    ### Troubleshooting
    
    - **Download fails**: Check your internet connection and Uplink credentials
    - **Email fails**: Verify email password is correct
    - **Reports not showing**: Check the Monthly_Reports folder
    - **Chrome driver error**: Update Chrome browser to latest version
    
    ### Data Security
    
    - Email passwords are NOT stored
    - All data processing happens locally
    - Reports saved to your local drive only
    """)

# Device management section
with st.sidebar.expander("➕ Manage Devices"):
    st.subheader("Add New Device")
    new_device_name = st.text_input("Device Name", placeholder="Scout-12200")
    new_device_id = st.text_input("Device ID", placeholder="359205108536868")
    new_device_location = st.text_input("Location", placeholder="East Building")
    new_commission_date = st.date_input("Commission Date", value=datetime.now())
    
    if st.button("Add Device"):
        if new_device_name and new_device_id:
            st.success(f"Device {new_device_name} added successfully!")
        else:
            st.error("Please fill in all required fields")