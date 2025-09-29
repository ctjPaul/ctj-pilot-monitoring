"""
CTJ Energy Pilot Monitoring Automation
Main orchestration module for automated report generation
"""

import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
import traceback

# Import supporting modules
from data_download import UplinkDownloader
from data_analyzer import PilotDataAnalyzer
from report_generator import PDFReportGenerator
from email_sender import EmailSender


def run_automation(config, status_callback=None):
    """
    Main automation function that orchestrates the entire report generation process.
    
    Args:
        config (dict): Configuration dictionary containing:
            - device: Device information
            - month: Month data with date ranges
            - email_recipients: List of email addresses
            - email_from: Sender email
            - email_password: Email password
            - send_email: Boolean for whether to send email
            - report_period_display: Display string for report period
        status_callback (function): Callback function for status updates
    
    Returns:
        dict: Result dictionary with success status, report path, and summary data
    """
    
    def update_status(message, progress):
        """Helper to update status if callback provided"""
        if status_callback:
            status_callback(message, progress)
        else:
            print(f"[{progress}%] {message}")
    
    try:
        # Create output directory
        output_dir = Path("Monthly_Reports")
        output_dir.mkdir(exist_ok=True)
        
        # Extract configuration
        device = config['device']
        month_data = config['month']
        device_name = device['name']
        device_id = device['id']
        
        # Generate filename
        report_filename = f"{device_name}_Monthly_Report_{config['report_period_display'].replace(' ', '_')}.pdf"
        report_path = output_dir / report_filename
        
        # Step 1: Login and download data from Uplink
        update_status("Step 1/5: Logging into Uplink system and downloading data...", 20)
        
        downloader = UplinkDownloader()
        csv_file = downloader.download_device_data(
            device_id=device_id,
            start_date=month_data['first_day'],
            end_date=month_data['last_day'],
            device_name=device_name
        )
        
        if not csv_file or not os.path.exists(csv_file):
            return {
                'success': False,
                'error': 'Failed to download data from Uplink system',
                'details': 'CSV file not found after download attempt'
            }
        
        # Step 2: Analyze the downloaded data
        update_status("Step 2/5: Analyzing pilot outage events...", 40)
        
        analyzer = PilotDataAnalyzer()
        analysis_results = analyzer.analyze_data(
            csv_file=csv_file,
            device=device,
            month_data=month_data
        )
        
        if not analysis_results['success']:
            return {
                'success': False,
                'error': 'Failed to analyze data',
                'details': analysis_results.get('error', 'Unknown error during analysis')
            }
        
        # Step 3: Generate PDF report
        update_status("Step 3/5: Generating PDF report...", 60)
        
        report_generator = PDFReportGenerator()
        pdf_result = report_generator.generate_report(
            analysis_results=analysis_results,
            device=device,
            month_data=month_data,
            output_path=report_path
        )
        
        if not pdf_result['success']:
            return {
                'success': False,
                'error': 'Failed to generate PDF report',
                'details': pdf_result.get('error', 'Unknown error during PDF generation')
            }
        
        # Step 4: Send email if requested
        email_sent = False
        if config['send_email'] and config['email_password']:
            update_status("Step 4/5: Sending email report...", 80)
            
            email_sender = EmailSender(
                from_email=config['email_from'],
                password=config['email_password']
            )
            
            email_result = email_sender.send_report(
                recipients=config['email_recipients'],
                device=device,
                month_period=config['report_period_display'],
                report_path=str(report_path),
                summary=analysis_results['summary']
            )
            
            email_sent = email_result['success']
            
            if not email_sent:
                # Don't fail entire process if email fails
                print(f"Warning: Email sending failed - {email_result.get('error', 'Unknown error')}")
        
        # Final status
        update_status("Step 5/5: Finalizing report...", 95)
        
        # Prepare summary for display
        summary = {
            'total_outages': analysis_results['summary']['total_outages'],
            'availability': f"{analysis_results['summary']['availability_percent']:.2f}%",
            'compliance': analysis_results['summary']['epa_compliance']
        }
        
        return {
            'success': True,
            'report_path': str(report_path),
            'email_sent': email_sent,
            'summary': summary,
            'detailed_results': analysis_results
        }
    
    except Exception as e:
        error_details = traceback.format_exc()
        return {
            'success': False,
            'error': str(e),
            'details': error_details
        }


def test_automation():
    """Test function for development and debugging"""
    
    test_config = {
        'device': {
            'id': '359205108536865',
            'name': 'Scout-12197',
            'commission_date': '06/10/2025',
            'location': 'Main Facility'
        },
        'month': {
            'display': 'September 2025',
            'month': 9,
            'year': 2025,
            'first_day': datetime(2025, 9, 1),
            'last_day': datetime(2025, 9, 30)
        },
        'email_recipients': ['test@example.com'],
        'email_from': 'reports@ctjenergy.com',
        'email_password': 'test_password',
        'send_email': False,
        'report_period_display': 'September 2025'
    }
    
    def print_status(message, progress):
        print(f"[{progress}%] {message}")
    
    result = run_automation(test_config, status_callback=print_status)
    
    print("\n" + "="*50)
    print("AUTOMATION RESULT:")
    print("="*50)
    print(f"Success: {result['success']}")
    
    if result['success']:
        print(f"Report Path: {result['report_path']}")
        print(f"Email Sent: {result['email_sent']}")
        print(f"Summary: {result['summary']}")
    else:
        print(f"Error: {result['error']}")
        if 'details' in result:
            print(f"Details:\n{result['details']}")


if __name__ == "__main__":
    # Run test when executed directly
    test_automation()