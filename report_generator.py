"""
PDF Report Generator Module
Creates professional PDF reports with charts and tables
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, 
    Spacer, PageBreak, Image, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.lineplots import LinePlot
from datetime import datetime
import os
from pathlib import Path


class PDFReportGenerator:
    """Generates professional PDF reports"""
    
    def __init__(self, logo_path=None):
        """
        Initialize report generator
        
        Args:
            logo_path (str): Path to company logo image
        """
        self.logo_path = logo_path or "ctj_logo.png"
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
    
    def _setup_custom_styles(self):
        """Setup custom paragraph styles"""
        
        # Title style
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#1f4788'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))
        
        # Section header style
        self.styles.add(ParagraphStyle(
            name='SectionHeader',
            parent=self.styles['Heading2'],
            fontSize=16,
            textColor=colors.HexColor('#1f4788'),
            spaceAfter=12,
            spaceBefore=12,
            fontName='Helvetica-Bold'
        ))
        
        # Subsection style
        self.styles.add(ParagraphStyle(
            name='SubSection',
            parent=self.styles['Heading3'],
            fontSize=12,
            textColor=colors.HexColor('#2c5aa0'),
            spaceAfter=6,
            fontName='Helvetica-Bold'
        ))
    
    def _create_header(self, device, month_data):
        """Create report header section"""
        
        elements = []
        
        # Logo (if exists)
        if os.path.exists(self.logo_path):
            try:
                img = Image(self.logo_path, width=2*inch, height=1*inch)
                img.hAlign = 'CENTER'
                elements.append(img)
                elements.append(Spacer(1, 0.2*inch))
            except:
                pass  # Skip logo if can't load
        
        # Title
        title = Paragraph(
            "Pilot Monitoring Monthly Report",
            self.styles['CustomTitle']
        )
        elements.append(title)
        
        # Device and period info
        info_text = f"""
        <b>Device:</b> {device['name']}<br/>
        <b>Device ID:</b> {device['id']}<br/>
        <b>Report Period:</b> {month_data['first_day'].strftime('%B %Y')}<br/>
        <b>Date Range:</b> {month_data['first_day'].strftime('%m/%d/%Y')} - {month_data['last_day'].strftime('%m/%d/%Y')}<br/>
        <b>Generated:</b> {datetime.now().strftime('%m/%d/%Y %I:%M %p')}
        """
        info = Paragraph(info_text, self.styles['Normal'])
        elements.append(info)
        elements.append(Spacer(1, 0.3*inch))
        
        return elements
    
    def _create_executive_summary(self, summary):
        """Create executive summary section"""
        
        elements = []
        
        # Section header
        header = Paragraph("Executive Summary", self.styles['SectionHeader'])
        elements.append(header)
        
        # Compliance status with color coding
        compliance = summary['compliance_details']
        status_color = colors.green if compliance['compliant'] else colors.red
        
        status_text = f"""
        <font color="{status_color.hexval()}"><b>EPA Compliance Status: {compliance['status']}</b></font>
        """
        status_para = Paragraph(status_text, self.styles['Normal'])
        elements.append(status_para)
        elements.append(Spacer(1, 0.2*inch))
        
        # Key metrics table
        metrics_data = [
            ['Metric', 'Value', 'Status'],
            ['Total Outages', str(summary['total_outages']), 
             '✓' if summary['total_outages'] <= 10 else '✗'],
            ['System Availability', f"{summary['availability_percent']:.2f}%",
             '✓' if summary['availability_percent'] >= 99.0 else '✗'],
            ['Total Outage Time', f"{summary['total_outage_minutes']:.2f} minutes", ''],
            ['Longest Outage', f"{summary['statistics']['max_duration_minutes']:.2f} minutes",
             '✓' if summary['statistics']['max_duration_minutes'] <= 4 else '✗'],
        ]
        
        metrics_table = Table(metrics_data, colWidths=[2.5*inch, 2*inch, 1*inch])
        metrics_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))
        
        elements.append(metrics_table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Compliance issues (if any)
        if not compliance['compliant'] and compliance['issues']:
            issues_header = Paragraph("Compliance Issues:", self.styles['SubSection'])
            elements.append(issues_header)
            
            for issue in compliance['issues']:
                issue_text = f"• {issue}"
                issue_para = Paragraph(issue_text, self.styles['Normal'])
                elements.append(issue_para)
            
            elements.append(Spacer(1, 0.2*inch))
        
        return elements
    
    def _create_outage_details(self, outages):
        """Create detailed outage table"""
        
        elements = []
        
        # Section header
        header = Paragraph("Outage Event Details", self.styles['SectionHeader'])
        elements.append(header)
        
        if not outages:
            no_outages = Paragraph("No outage events recorded this period.", self.styles['Normal'])
            elements.append(no_outages)
            return elements
        
        # Create outage table
        table_data = [['Event #', 'Start Time', 'End Time', 'Duration (minutes)', 'Status']]
        
        for idx, outage in enumerate(outages, 1):
            status = '⚠️ Ongoing' if outage.get('ongoing', False) else '✓ Resolved'
            duration_status = '✗ Exceeds limit' if outage['duration_minutes'] > 240 else '✓ Within limit'
            
            table_data.append([
                str(idx),
                outage['start'].strftime('%m/%d/%Y %H:%M'),
                outage['end'].strftime('%m/%d/%Y %H:%M'),
                f"{outage['duration_minutes']:.2f}",
                duration_status
            ])
        
        outage_table = Table(table_data, colWidths=[0.7*inch, 1.8*inch, 1.8*inch, 1.3*inch, 1.4*inch])
        outage_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))
        
        elements.append(outage_table)
        elements.append(Spacer(1, 0.3*inch))
        
        return elements
    
    def _create_statistics_section(self, stats):
        """Create statistical analysis section"""
        
        elements = []
        
        # Section header
        header = Paragraph("Statistical Analysis", self.styles['SectionHeader'])
        elements.append(header)
        
        # Statistics text
        stats_text = f"""
        <b>Mean Outage Duration:</b> {stats['mean_duration_minutes']:.2f} minutes<br/>
        <b>Median Outage Duration:</b> {stats['median_duration_minutes']:.2f} minutes<br/>
        <b>Maximum Outage Duration:</b> {stats['max_duration_minutes']:.2f} minutes<br/>
        <b>Minimum Outage Duration:</b> {stats['min_duration_minutes']:.2f} minutes<br/>
        <b>Standard Deviation:</b> {stats['std_duration_minutes']:.2f} minutes
        """
        
        stats_para = Paragraph(stats_text, self.styles['Normal'])
        elements.append(stats_para)
        elements.append(Spacer(1, 0.3*inch))
        
        return elements
    
    def _create_footer(self):
        """Create report footer"""
        
        elements = []
        
        elements.append(Spacer(1, 0.5*inch))
        
        footer_text = """
        <para align=center>
        <font size=8>
        This report is generated automatically by the CTJ Energy Pilot Monitoring System.<br/>
        For questions or concerns, please contact your CTJ Energy representative.<br/>
        <b>CTJ Energy Solutions</b> | www.ctjenergy.com | support@ctjenergy.com
        </font>
        </para>
        """
        
        footer = Paragraph(footer_text, self.styles['Normal'])
        elements.append(footer)
        
        return elements
    
    def generate_report(self, analysis_results, device, month_data, output_path):
        """
        Generate the complete PDF report
        
        Args:
            analysis_results (dict): Results from data analysis
            device (dict): Device information
            month_data (dict): Month information
            output_path (str): Output file path
        
        Returns:
            dict: Result with success status and path
        """
        
        try:
            # Create document
            doc = SimpleDocTemplate(
                str(output_path),
                pagesize=letter,
                rightMargin=0.75*inch,
                leftMargin=0.75*inch,
                topMargin=0.75*inch,
                bottomMargin=0.75*inch
            )
            
            # Build content
            story = []
            
            # Add sections
            story.extend(self._create_header(device, month_data))
            story.extend(self._create_executive_summary(analysis_results['summary']))
            story.extend(self._create_outage_details(analysis_results['outages']))
            story.extend(self._create_statistics_section(analysis_results['summary']['statistics']))
            story.extend(self._create_footer())
            
            # Build PDF
            doc.build(story)
            
            print(f"✓ PDF report generated: {output_path}")
            
            return {
                'success': True,
                'path': str(output_path)
            }
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"✗ Failed to generate PDF: {e}")
            print(error_details)
            
            return {
                'success': False,
                'error': str(e),
                'details': error_details
            }