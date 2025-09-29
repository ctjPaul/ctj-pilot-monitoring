"""
Email Sender Module
Handles sending email reports with attachments
"""

import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
from pathlib import Path


class EmailSender:
    """Handles sending email reports"""
    
    def __init__(self, from_email, password, smtp_server=None, smtp_port=None):
        """
        Initialize email sender
        
        Args:
            from_email (str): Sender email address
            password (str): Email password
            smtp_server (str): SMTP server address (auto-detect if None)
            smtp_port (int): SMTP port (auto-detect if None)
        """
        self.from_email = from_email
        self.password = password
        
        # Auto-detect SMTP settings based on email provider
        if smtp_server is None or smtp_port is None:
            self.smtp_server, self.smtp_port = self._detect_smtp_settings(from_email)
        else:
            self.smtp_server = smtp_server
            self.smtp_port = smtp_port
    
    def _detect_smtp_settings(self, email):
        """
        Auto-detect SMTP settings based on email domain
        
        Args:
            email (str): Email address
        
        Returns:
            tuple: (smtp_server, smtp_port)
        """
        domain = email.split('@')[1].lower()
        
        smtp_configs = {
            'gmail.com': ('smtp.gmail.com', 587),
            'outlook.com': ('smtp-mail.outlook.com', 587),
            'hotmail.com': ('smtp-mail.outlook.com', 587),
            'live.com': ('smtp-mail.outlook.com', 587),
            'office365.com': ('smtp.office365.com', 587),
            'yahoo.com': ('smtp.mail.yahoo.com', 587),
            'ctjenergy.com': ('smtp.office365.com', 587),
        }
        
        return smtp_configs.get(domain, ('smtp.gmail.com', 587))
    
    def _create_email_body(self, device, month_period, summary):
        """
        Create HTML email body
        
        Args:
            device (dict): Device information
            month_period (str): Report period display string
            summary (dict): Analysis summary data
        
        Returns:
            str: HTML email body
        """
        
        # Color for compliance status
        compliance_color = "green" if summary['compliance_details']['compliant'] else "red"
        
        html_body = f"""
        <html>
        <head>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    color: #333;
                }}
                .header {{
                    background-color: #1f4788;
                    color: white;
                    padding: 20px;
                    text-align: center;
                }}
                .content {{
                    padding: 20px;
                }}
                .summary-box {{
                    background-color: #f4f4f4;
                    border-left: 4px solid #1f4788;
                    padding: 15px;
                    margin: 20px 0;
                }}
                .compliance {{
                    font-size: 18px;
                    font-weight: bold;
                    color: {compliance_color};
                    margin: 15px 0;
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    margin: 15px 0;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    padding: 12px;
                    text-align: left;
                }}
                th {{
                    background-color: #1f4788;
                    color: white;
                }}
                tr:nth-child(even) {{
                    background-color: #f9f9f9;
                }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>CTJ Energy Pilot Monitoring Report</h1>
            </div>
            
            <div class="content">
                <h2>Monthly Report - {month_period}</h2>
                
                <p>Dear Team,</p>
                
                <p>Please find attached the automated pilot monitoring report for <strong>{device['name']}</strong> 
                covering the period of <strong>{month_period}</strong>.</p>
                
                <div class="summary-box">
                    <h3>Executive Summary</h3>
                    
                    <div class="compliance">
                        EPA Compliance Status: {summary['compliance_details']['status']}
                    </div>
                    
                    <table>
                        <tr>
                            <th>Metric</th>
                            <th>Value</th>
                        </tr>
                        <tr>
                            <td>Device</td>
                            <td>{device['name']} ({device['id']})</td>
                        </tr>
                        <tr>
                            <td>Total Outage Events</td>
                            <td>{summary['total_outages']}</td>
                        </tr>
                        <tr>
                            <td>System Availability</td>
                            <td>{summary['availability_percent']:.2f}%</td>
                        </tr>
                        <tr>
                            <td>Total Downtime</td>
                            <td>{summary['total_outage_minutes']:.2f} minutes</td>
                        </tr>
                        <tr>
                            <td>Longest Outage</td>
                            <td>{summary['statistics']['max_duration_minutes']:.2f} minutes</td>
                        </tr>
                        <tr>
                            <td>Average Outage Duration</td>
                            <td>{summary['statistics']['mean_duration_minutes']:.2f} minutes</td>
                        </tr>
                    </table>
                </div>
        """
        
        # Add compliance issues if any
        if not summary['compliance_details']['compliant']:
            html_body += """
                <div style="background-color: #fff3cd; border-left: 4px solid #ff9800; padding: 15px; margin: 20px 0;">
                    <h3 style="color: #ff9800; margin-top: 0;">Compliance Issues Detected</h3>
                    <ul>
            """
            for issue in summary['compliance_details']['issues']:
                html_body += f"<li>{issue}</li>"
            html_body += """
                    </ul>
                </div>
            """
        
        html_body += f"""
                <h3>Next Steps</h3>
                <ul>
                    <li>Review the detailed PDF report attached to this email</li>
                    <li>Address any compliance issues identified</li>
                    <li>Schedule maintenance if required</li>
                    <li>Contact CTJ Energy support if you have any questions</li>
                </ul>
                
                <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; color: #666;">
                    <p><strong>Report Generated:</strong> {datetime.now().strftime('%m/%d/%Y %I:%M %p')}</p>
                    <p><strong>CTJ Energy Solutions</strong><br/>
                    Automated Pilot Monitoring System v2.0<br/>
                    For support, contact: support@ctjenergy.com</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        return html_body
    
    def send_report(self, recipients, device, month_period, report_path, summary):
        """
        Send email report with PDF attachment
        
        Args:
            recipients (list): List of recipient email addresses
            device (dict): Device information
            month_period (str): Report period display string
            report_path (str): Path to PDF report file
            summary (dict): Analysis summary data
        
        Returns:
            dict: Result with success status
        """
        
        try:
            # Create message
            msg = MIMEMultipart('alternative')
            msg['From'] = self.from_email
            msg['To'] = ', '.join(recipients)
            msg['Subject'] = f"CTJ Energy Pilot Monitoring Report - {device['name']} - {month_period}"
            
            # Create plain text version
            text_body = f"""
            CTJ Energy Pilot Monitoring Report
            
            Device: {device['name']} ({device['id']})
            Report Period: {month_period}
            
            EPA Compliance: {summary['compliance_details']['status']}
            Total Outages: {summary['total_outages']}
            Availability: {summary['availability_percent']:.2f}%
            
            Please see the attached PDF for the full report.
            
            ---
            CTJ Energy Solutions
            Automated Pilot Monitoring System
            """
            
            # Create HTML version
            html_body = self._create_email_body(device, month_period, summary)
            
            # Attach both versions
            part1 = MIMEText(text_body, 'plain')
            part2 = MIMEText(html_body, 'html')
            msg.attach(part1)
            msg.attach(part2)
            
            # Attach PDF report
            if os.path.exists(report_path):
                with open(report_path, 'rb') as f:
                    pdf_attachment = MIMEApplication(f.read(), _subtype='pdf')
                    pdf_attachment.add_header(
                        'Content-Disposition', 
                        'attachment', 
                        filename=os.path.basename(report_path)
                    )
                    msg.attach(pdf_attachment)
                
                print(f"✓ Attached PDF: {os.path.basename(report_path)}")
            else:
                print("⚠️ Warning: PDF file not found, sending email without attachment")
            
            # Connect to SMTP server and send
            print(f"Connecting to {self.smtp_server}:{self.smtp_port}...")
            
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                print("✓ Connected to SMTP server")
                
                # Login
                server.login(self.from_email, self.password)
                print("✓ Logged in successfully")
                
                # Send email
                server.send_message(msg)
                print(f"✓ Email sent to {len(recipients)} recipient(s)")
            
            return {
                'success': True,
                'recipients': recipients,
                'message': f"Email sent successfully to {len(recipients)} recipient(s)"
            }
            
        except smtplib.SMTPAuthenticationError:
            error_msg = "Authentication failed. Please check email credentials."
            print(f"✗ {error_msg}")
            return {
                'success': False,
                'error': error_msg,
                'details': 'SMTP authentication error - verify email and password'
            }
            
        except smtplib.SMTPException as e:
            error_msg = f"SMTP error: {str(e)}"
            print(f"✗ {error_msg}")
            return {
                'success': False,
                'error': error_msg,
                'details': str(e)
            }
            
        except Exception as e:
            error_msg = f"Failed to send email: {str(e)}"
            print(f"✗ {error_msg}")
            import traceback
            return {
                'success': False,
                'error': error_msg,
                'details': traceback.format_exc()
            }