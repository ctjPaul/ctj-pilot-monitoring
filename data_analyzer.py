"""
Pilot Data Analyzer Module
Analyzes pilot monitoring data from CSV files
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path


class PilotDataAnalyzer:
    """Analyzes pilot monitoring data from CSV files"""
    
    def __init__(self):
        """Initialize the analyzer"""
        # EPA compliance thresholds
        self.MAX_OUTAGES_PER_MONTH = 10
        self.MAX_OUTAGE_DURATION_MINUTES = 60
        self.MIN_AVAILABILITY_PERCENT = 99.0
    
    def load_data(self, csv_file):
        """
        Load CSV data file
        
        Args:
            csv_file (str): Path to CSV file
        
        Returns:
            pd.DataFrame: Loaded data or None if failed
        """
        try:
            # Try different encodings
            for encoding in ['utf-8', 'latin-1', 'iso-8859-1']:
                try:
                    df = pd.read_csv(csv_file, encoding=encoding)
                    print(f"✓ Successfully loaded CSV with {len(df)} rows")
                    return df
                except UnicodeDecodeError:
                    continue
            
            print("✗ Could not load CSV with any encoding")
            return None
            
        except Exception as e:
            print(f"✗ Error loading CSV: {e}")
            return None
    
    def parse_timestamps(self, df, timestamp_col='Time'):
        """
        Parse timestamp column to datetime
        
        Args:
            df (pd.DataFrame): Input dataframe
            timestamp_col (str): Name of timestamp column
        
        Returns:
            pd.DataFrame: DataFrame with parsed timestamps
        """
        try:
            # The column is already called 'Time' in the Uplink data
            if timestamp_col in df.columns:
                df[timestamp_col] = pd.to_datetime(df[timestamp_col])
                print(f"✓ Parsed timestamps from column: {timestamp_col}")
                return df
            else:
                print(f"✗ Column '{timestamp_col}' not found")
                return df
            
        except Exception as e:
            print(f"✗ Error parsing timestamps: {e}")
            return df
    
    def identify_outages(self, df, message_col='Message', time_col='Time'):
        """
        Identify pilot outage events from Uplink data
        
        Args:
            df (pd.DataFrame): Input dataframe with Time and Message columns
            message_col (str): Name of message column
            time_col (str): Name of time column
        
        Returns:
            list: List of outage event dictionaries
        """
        outages = []
        in_outage = False
        outage_start = None
        
        print(f"Analyzing {len(df)} rows for pilot status changes...")
        
        for idx, row in df.iterrows():
            message = str(row[message_col]).strip()
            timestamp = row[time_col]
            
            # Check if this is a pilot status message
            if 'Pilot Inactive' in message:
                # Start of outage
                if not in_outage:
                    in_outage = True
                    outage_start = timestamp
                    print(f"  Outage started: {timestamp}")
                    
            elif 'Pilot Active' in message:
                # End of outage
                if in_outage:
                    in_outage = False
                    outage_end = timestamp
                    
                    # Calculate duration
                    duration = outage_end - outage_start
                    duration_minutes = duration.total_seconds() / 60
                    
                    outages.append({
                        'start': outage_start,
                        'end': outage_end,
                        'duration': duration,
                        'duration_minutes': duration_minutes
                    })
                    print(f"  Outage ended: {timestamp} (Duration: {duration_minutes:.2f} minutes)")
        
        # Handle case where month ends with an ongoing outage
        if in_outage and outage_start:
            outage_end = df[time_col].iloc[-1]
            duration = outage_end - outage_start
            duration_minutes = duration.total_seconds() / 60
            
            outages.append({
                'start': outage_start,
                'end': outage_end,
                'duration': duration,
                'duration_minutes': duration_minutes,
                'ongoing': True
            })
            print(f"  Ongoing outage at end of period (Duration: {duration_minutes:.2f} minutes)")
        
        print(f"✓ Identified {len(outages)} outage events")
        return outages
    
    def calculate_availability(self, df, outages, month_data):
        """
        Calculate pilot availability percentage
        
        Args:
            df (pd.DataFrame): Data frame
            outages (list): List of outage events
            month_data (dict): Month information with first_day and last_day
        
        Returns:
            float: Availability percentage
        """
        try:
            # Calculate total minutes in month
            total_minutes = (month_data['last_day'] - month_data['first_day']).total_seconds() / 60
            
            # Calculate total outage minutes
            total_outage_minutes = sum(o['duration_minutes'] for o in outages)
            
            # Calculate availability
            availability_percent = ((total_minutes - total_outage_minutes) / total_minutes) * 100
            
            print(f"✓ Calculated availability: {availability_percent:.2f}%")
            return availability_percent
            
        except Exception as e:
            print(f"✗ Error calculating availability: {e}")
            return 0.0
    
    def check_epa_compliance(self, outages, availability_percent):
        """
        Check EPA compliance status
        
        Args:
            outages (list): List of outage events
            availability_percent (float): Availability percentage
        
        Returns:
            dict: Compliance status and details
        """
        compliance_issues = []
        
        # Check number of outages
        if len(outages) > self.MAX_OUTAGES_PER_MONTH:
            compliance_issues.append(
                f"Exceeded maximum outages: {len(outages)} > {self.MAX_OUTAGES_PER_MONTH}"
            )
        
        # Check individual outage durations
        long_outages = [o for o in outages if o['duration_minutes'] > self.MAX_OUTAGE_DURATION_MINUTES]
        if long_outages:
            compliance_issues.append(
                f"Found {len(long_outages)} outage(s) exceeding {self.MAX_OUTAGE_DURATION_MINUTES} minutes"
            )
        
        # Check overall availability
        if availability_percent < self.MIN_AVAILABILITY_PERCENT:
            compliance_issues.append(
                f"Availability below minimum: {availability_percent:.2f}% < {self.MIN_AVAILABILITY_PERCENT}%"
            )
        
        is_compliant = len(compliance_issues) == 0
        
        status = "COMPLIANT" if is_compliant else "NON-COMPLIANT"
        
        print(f"✓ EPA Compliance: {status}")
        
        return {
            'compliant': is_compliant,
            'status': status,
            'issues': compliance_issues
        }
    
    def calculate_statistics(self, outages):
        """
        Calculate statistical metrics for outages
        
        Args:
            outages (list): List of outage events
        
        Returns:
            dict: Statistical metrics
        """
        if not outages:
            return {
                'mean_duration_minutes': 0,
                'median_duration_minutes': 0,
                'max_duration_minutes': 0,
                'min_duration_minutes': 0,
                'std_duration_minutes': 0
            }
        
        durations = [o['duration_minutes'] for o in outages]
        
        return {
            'mean_duration_minutes': np.mean(durations),
            'median_duration_minutes': np.median(durations),
            'max_duration_minutes': np.max(durations),
            'min_duration_minutes': np.min(durations),
            'std_duration_minutes': np.std(durations)
        }
    
    def analyze_data(self, csv_file, device, month_data):
        """
        Main analysis method - UPDATED FOR UPLINK DATA FORMAT
        
        Args:
            csv_file (str): Path to Excel file (converted from XML)
            device (dict): Device information
            month_data (dict): Month information
        
        Returns:
            dict: Complete analysis results
        """
        try:
            # Load data (it's actually an Excel file now, not CSV)
            print(f"Loading data from: {csv_file}")
            
            if csv_file.endswith('.xlsx') or csv_file.endswith('.xls'):
                df = pd.read_excel(csv_file)
            else:
                df = self.load_data(csv_file)
            
            if df is None:
                return {'success': False, 'error': 'Failed to load data file'}
            
            print(f"✓ Loaded {len(df)} rows")
            print(f"Columns: {list(df.columns)}")
            
            # Check for required columns - UPLINK FORMAT
            if 'Time' not in df.columns or 'Message' not in df.columns:
                return {
                    'success': False, 
                    'error': f'Required columns not found. Expected "Time" and "Message", got: {list(df.columns)}'
                }
            
            # Parse timestamps
            df = self.parse_timestamps(df, timestamp_col='Time')
            
            # Sort by timestamp
            df = df.sort_values('Time')
            
            # Filter data to month range
            print(f"Filtering data to date range: {month_data['first_day']} to {month_data['last_day']}")
            df = df[
                (df['Time'] >= month_data['first_day']) &
                (df['Time'] <= month_data['last_day'])
            ]
            
            print(f"✓ Filtered to {len(df)} rows within date range")
            
            # Filter to only pilot status messages
            pilot_messages = df[df['Message'].str.contains('Pilot', case=False, na=False)]
            print(f"✓ Found {len(pilot_messages)} pilot status messages")
            
            # Identify outages using the Message column
            outages = self.identify_outages(df, message_col='Message', time_col='Time')
            
            # Calculate availability
            availability_percent = self.calculate_availability(df, outages, month_data)
            
            # Check EPA compliance
            compliance = self.check_epa_compliance(outages, availability_percent)
            
            # Calculate statistics
            stats = self.calculate_statistics(outages)
            
            # Prepare summary
            summary = {
                'total_outages': len(outages),
                'total_outage_minutes': sum(o['duration_minutes'] for o in outages),
                'availability_percent': availability_percent,
                'epa_compliance': compliance['status'],
                'compliance_details': compliance,
                'statistics': stats
            }
            
            print("\n" + "="*50)
            print("ANALYSIS SUMMARY")
            print("="*50)
            print(f"Total Outages: {summary['total_outages']}")
            print(f"Total Downtime: {summary['total_outage_minutes']:.2f} minutes")
            print(f"Availability: {summary['availability_percent']:.2f}%")
            print(f"EPA Compliance: {summary['epa_compliance']}")
            print("="*50)
            
            return {
                'success': True,
                'data': df,
                'outages': outages,
                'summary': summary,
                'device': device,
                'month_data': month_data
            }
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"✗ Analysis failed: {e}")
            print(error_details)
            return {
                'success': False,
                'error': str(e),
                'details': error_details
            }