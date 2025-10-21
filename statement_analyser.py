#!/usr/bin/env python3
"""
Statement Analyser for PhonePe and Google Pay

This script analyzes transaction statements from PhonePe and Google Pay PDF files.
It automatically detects the PDF type, parses transactions, performs comprehensive
financial analysis, generates visualizations, and exports everything to Excel.

Features:
- Automatic PDF type detection (PhonePe/Google Pay)
- Password-protected PDF support
- Single month and multi-month analysis modes
- Comprehensive financial statistics and insights
- Interactive visualizations
- Excel export with embedded charts

Author: Statement Analyser
Version: 2.0
"""

import pdfplumber
import getpass
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import re
import warnings
import logging
import sys
import os
from io import BytesIO

warnings.filterwarnings("ignore", category=UserWarning)

logging.getLogger("pdfminer.pdfinterp").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfparser").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfdocument").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfpage").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfpagecache").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdfdevice").setLevel(logging.ERROR)
logging.getLogger("pdfminer.layout").setLevel(logging.ERROR)


class SingleMonthAnalysis:
    """
    Analyzes transactions for a single month (last 30 days).
    
    Provides detailed analysis including:
    - Summary statistics (totals, averages, medians)
    - Top merchants analysis
    - Spending categories
    - Weekday and time-of-day patterns
    - Transaction frequency metrics
    - Savings insights
    """
    
    def __init__(self, df):
        """
        Initialize single month analysis.
        
        Args:
            df: DataFrame containing transaction data
        """
        last_30_days = datetime.now() - timedelta(days=30)
        self.df = df[df['Date'] >= last_30_days].copy()
        self.summary_data = {}
        self.plots = {}

    def summary_stats(self):
        """Calculate and display comprehensive summary statistics for the last 30 days."""
        df = self.df
        debit_df = df[df['Type']=='Debit']
        credit_df = df[df['Type']=='Credit']
        
        total_debit = debit_df['Amount'].sum()
        total_credit = credit_df['Amount'].sum()
        avg_debit = debit_df['Amount'].mean()
        avg_credit = credit_df['Amount'].mean()
        median_debit = debit_df['Amount'].median()
        median_credit = credit_df['Amount'].median()
        
        max_txn = df.loc[df['Amount'].idxmax()]
        min_debit = debit_df['Amount'].min() if len(debit_df) > 0 else 0
        min_credit = credit_df['Amount'].min() if len(credit_df) > 0 else 0
        
        net_flow = total_credit - total_debit
        debit_count = len(debit_df)
        credit_count = len(credit_df)
        total_txn = len(df)
        
        avg_daily_spend = total_debit / 30 if total_debit > 0 else 0
        
        start_date = df['Date'].min()
        end_date = df['Date'].max()
        days_covered = (end_date - start_date).days + 1
        
        self.summary_data['summary'] = pd.DataFrame({
            'Metric': [
                'Analysis Period',
                'Days Covered',
                '─' * 50,
                'Total Debit', 'Total Credit', 'Net Cash Flow',
                '─' * 50,
                'Average Debit', 'Average Credit',
                'Median Debit', 'Median Credit',
                'Minimum Debit', 'Minimum Credit',
                '─' * 50,
                'Largest Transaction',
                '─' * 50,
                'Total Transactions', 'Debit Transactions', 'Credit Transactions',
                'Average Daily Spending'
            ],
            'Value': [
                f"{start_date.date()} to {end_date.date()}",
                str(days_covered),
                '',
                f"₹{total_debit:,.2f}",
                f"₹{total_credit:,.2f}",
                f"₹{net_flow:,.2f}" + (" (Surplus)" if net_flow > 0 else " (Deficit)" if net_flow < 0 else " (Balanced)"),
                '',
                f"₹{avg_debit:,.2f}",
                f"₹{avg_credit:,.2f}",
                f"₹{median_debit:,.2f}",
                f"₹{median_credit:,.2f}",
                f"₹{min_debit:,.2f}",
                f"₹{min_credit:,.2f}",
                '',
                f"₹{max_txn['Amount']:.2f} at {max_txn['Merchant']} on {max_txn['Date'].date()}",
                '',
                str(total_txn),
                str(debit_count),
                str(credit_count),
                f"₹{avg_daily_spend:,.2f}"
            ]
        })
        
        print("=" * 70)
        print("LAST 30 DAYS SUMMARY")
        print("=" * 70)
        print(f"\nAnalysis Period: {start_date.date()} to {end_date.date()}")
        print(f"Days Covered: {days_covered}")
        print(f"\n{'─' * 70}")
        print(f"Total Debit:  ₹{total_debit:>15,.2f}")
        print(f"Total Credit: ₹{total_credit:>15,.2f}")
        print(f"Net Flow:     ₹{net_flow:>15,.2f}" + (" (Surplus)" if net_flow > 0 else " (Deficit)" if net_flow < 0 else " (Balanced)"))
        print(f"{'─' * 70}")
        print(f"Average Debit:  ₹{avg_debit:>13,.2f} | Median: ₹{median_debit:,.2f}")
        print(f"Average Credit: ₹{avg_credit:>13,.2f} | Median: ₹{median_credit:,.2f}")
        print(f"{'─' * 70}")
        print(f"Average Daily Spending: ₹{avg_daily_spend:,.2f}")
        print(f"{'─' * 70}")
        print(f"Total Transactions: {total_txn} (Debit: {debit_count}, Credit: {credit_count})")
        print(f"\nLargest Transaction:")
        print(f"  Amount:   ₹{max_txn['Amount']:,.2f}")
        print(f"  Merchant: {max_txn['Merchant']}")
        print(f"  Date:     {max_txn['Date'].date()}")
        print("=" * 70)

    def top_merchants(self, n=10):
        """Identify and display top N merchants by total spending."""
        df = self.df
        top = (df[df['Type']=='Debit']
               .groupby('Merchant')['Amount']
               .agg(['sum', 'count', 'mean'])
               .sort_values('sum', ascending=False)
               .head(n))
        
        top_df = top.reset_index()
        top_df.columns = ['Merchant', 'Total Spent (INR)', 'Transactions', 'Average (INR)']
        top_df['Total Spent (INR)'] = top_df['Total Spent (INR)'].round(2)
        top_df['Average (INR)'] = top_df['Average (INR)'].round(2)
        
        self.summary_data['top_merchants'] = top_df
        
        print("\n" + "=" * 70)
        print(f"TOP {n} MERCHANTS (LAST 30 DAYS)")
        print("=" * 70)
        for i, row in top_df.iterrows():
            print(f"{i+1:2d}. {row['Merchant']:<35} ₹{row['Total Spent (INR)']:>12,.2f} ({int(row['Transactions'])} txns)")
        print("=" * 70)

    def plot_daily_spend(self):
        """Generate and display a bar chart of daily spending."""
        df = self.df
        daily_spend = df[df['Type']=='Debit'].groupby(df['Date'].dt.date)['Amount'].sum()
        
        fig = plt.figure(figsize=(10,5))
        daily_spend.plot(kind='bar', title='Daily Spending in Last 30 Days', ylabel='Amount (INR)')
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['daily_spend'] = buf
        
        plt.savefig('daily_spend_last_30_days.png')
        print("\nSaved plot: daily_spend_last_30_days.png")
        plt.show()

    def plot_debit_vs_credit(self):
        """Generate and display a bar chart comparing total debit vs credit."""
        df = self.df
        type_summary = df.groupby('Type')['Amount'].sum()
        
        fig = plt.figure(figsize=(5,5))
        type_summary.plot(kind='bar', title='Debit vs Credit in Last 30 Days', ylabel='Amount (INR)')
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['debit_vs_credit'] = buf
        
        plt.savefig('debit_vs_credit_last_30_days.png')
        print("Saved plot: debit_vs_credit_last_30_days.png")
        plt.show()

    def spending_categories(self):
        """
        Categorize spending into predefined ranges and analyze patterns.
        Categories: Under ₹100, ₹100-500, ₹500-1000, ₹1000-5000, Above ₹5000
        """
        df = self.df[self.df['Type']=='Debit']
        
        categories = []
        for _, row in df.iterrows():
            amount = row['Amount']
            if amount < 100:
                cat = 'Under ₹100'
            elif amount < 500:
                cat = '₹100-500'
            elif amount < 1000:
                cat = '₹500-1000'
            elif amount < 5000:
                cat = '₹1000-5000'
            else:
                cat = 'Above ₹5000'
            categories.append(cat)
        
        df_cat = df.copy()
        df_cat['Category'] = categories
        
        category_summary = df_cat.groupby('Category').agg({
            'Amount': ['sum', 'count', 'mean']
        }).round(2)
        category_summary.columns = ['Total Amount (INR)', 'Transaction Count', 'Average (INR)']
        category_summary = category_summary.reset_index()
        
        category_order = ['Under ₹100', '₹100-500', '₹500-1000', '₹1000-5000', 'Above ₹5000']
        category_summary['Category'] = pd.Categorical(category_summary['Category'], categories=category_order, ordered=True)
        category_summary = category_summary.sort_values('Category')
        
        self.summary_data['spending_categories'] = category_summary
        
        print("\n" + "=" * 70)
        print("SPENDING BY CATEGORY")
        print("=" * 70)
        print(category_summary.to_string(index=False))
        print("=" * 70)
        
        # Plot
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        
        category_summary.plot(x='Category', y='Total Amount (INR)', kind='bar', ax=ax1, legend=False, color='steelblue')
        ax1.set_title('Total Spending by Category')
        ax1.set_ylabel('Amount (INR)')
        ax1.tick_params(axis='x', rotation=45)
        
        category_summary.plot(x='Category', y='Transaction Count', kind='bar', ax=ax2, legend=False, color='coral')
        ax2.set_title('Transaction Count by Category')
        ax2.set_ylabel('Number of Transactions')
        ax2.tick_params(axis='x', rotation=45)
        
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['spending_categories'] = buf
        
        plt.savefig('spending_categories.png')
        print("Saved plot: spending_categories.png")
        plt.show()

    def weekday_analysis(self):
        """
        Analyze spending patterns by day of week.
        Identifies highest and lowest spending days.
        """
        df = self.df.copy()
        df['Weekday'] = df['Date'].dt.day_name()
        
        weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        
        debit_by_weekday = df[df['Type']=='Debit'].groupby('Weekday').agg({
            'Amount': ['sum', 'count', 'mean']
        }).round(2)
        debit_by_weekday.columns = ['Total Spent (INR)', 'Transaction Count', 'Average (INR)']
        debit_by_weekday = debit_by_weekday.reset_index()
        debit_by_weekday['Weekday'] = pd.Categorical(debit_by_weekday['Weekday'], categories=weekday_order, ordered=True)
        debit_by_weekday = debit_by_weekday.sort_values('Weekday')
        
        self.summary_data['weekday_spending'] = debit_by_weekday
        
        print("\n" + "=" * 70)
        print("SPENDING BY DAY OF WEEK")
        print("=" * 70)
        print(debit_by_weekday.to_string(index=False))
        
        # Find highest and lowest spending days
        max_day = debit_by_weekday.loc[debit_by_weekday['Total Spent (INR)'].idxmax()]
        min_day = debit_by_weekday.loc[debit_by_weekday['Total Spent (INR)'].idxmin()]
        print(f"\n{'─' * 70}")
        print(f"📈 Highest spending day: {max_day['Weekday']} (₹{max_day['Total Spent (INR)']:,.2f})")
        print(f"📉 Lowest spending day:  {min_day['Weekday']} (₹{min_day['Total Spent (INR)']:,.2f})")
        print("=" * 70)
        
        # Plot
        fig = plt.figure(figsize=(10, 5))
        plt.bar(debit_by_weekday['Weekday'], debit_by_weekday['Total Spent (INR)'], color='teal')
        plt.title('Spending by Day of Week')
        plt.xlabel('Day')
        plt.ylabel('Total Amount (INR)')
        plt.xticks(rotation=45)
        plt.grid(axis='y', alpha=0.3)
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['weekday_spending'] = buf
        
        plt.savefig('weekday_spending.png')
        print("Saved plot: weekday_spending.png")
        plt.show()

    def time_of_day_analysis(self):
        """
        Analyze spending patterns by time of day.
        Categories: Morning (5AM-12PM), Afternoon (12PM-5PM), Evening (5PM-9PM), Night (9PM-5AM)
        """
        df = self.df[self.df['Type']=='Debit'].copy()
        
        if 'Time' not in df.columns or df['Time'].isna().all():
            print("\n" + "=" * 70)
            print("TIME OF DAY ANALYSIS")
            print("=" * 70)
            print("⚠️  Time data not available in transactions")
            print("=" * 70)
            return
        
        def categorize_time(time_str):
            if pd.isna(time_str) or time_str == '':
                return 'Unknown'
            try:
                time_str = str(time_str).strip()
                if 'AM' in time_str or 'PM' in time_str:
                    time_obj = datetime.strptime(time_str, '%I:%M %p')
                else:
                    time_obj = datetime.strptime(time_str, '%H:%M')
                
                hour = time_obj.hour
                if 5 <= hour < 12:
                    return 'Morning (5AM-12PM)'
                elif 12 <= hour < 17:
                    return 'Afternoon (12PM-5PM)'
                elif 17 <= hour < 21:
                    return 'Evening (5PM-9PM)'
                else:
                    return 'Night (9PM-5AM)'
            except:
                return 'Unknown'
        
        df['TimeOfDay'] = df['Time'].apply(categorize_time)
        
        time_summary = df.groupby('TimeOfDay').agg({
            'Amount': ['sum', 'count', 'mean']
        }).round(2)
        time_summary.columns = ['Total Spent (INR)', 'Transaction Count', 'Average (INR)']
        time_summary = time_summary.reset_index()
        
        time_order = ['Morning (5AM-12PM)', 'Afternoon (12PM-5PM)', 'Evening (5PM-9PM)', 'Night (9PM-5AM)', 'Unknown']
        time_summary['TimeOfDay'] = pd.Categorical(time_summary['TimeOfDay'], categories=time_order, ordered=True)
        time_summary = time_summary.sort_values('TimeOfDay')
        
        self.summary_data['time_of_day_spending'] = time_summary
        
        print("\n" + "=" * 70)
        print("SPENDING BY TIME OF DAY")
        print("=" * 70)
        print(time_summary.to_string(index=False))
        print("=" * 70)
        
        # Plot
        fig = plt.figure(figsize=(10, 5))
        colors = ['#FFD700', '#FF8C00', '#FF6347', '#4B0082']
        plt.bar(time_summary['TimeOfDay'], time_summary['Total Spent (INR)'], color=colors[:len(time_summary)])
        plt.title('Spending by Time of Day')
        plt.xlabel('Time Period')
        plt.ylabel('Total Amount (INR)')
        plt.xticks(rotation=45, ha='right')
        plt.grid(axis='y', alpha=0.3)
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['time_of_day_spending'] = buf
        
        plt.savefig('time_of_day_spending.png')
        print("Saved plot: time_of_day_spending.png")
        plt.show()

    def transaction_frequency(self):
        """
        Analyze transaction frequency patterns.
        Calculates average, max, min transactions per day and identifies most active day.
        """
        df = self.df.copy()
        
        # Transactions per day
        daily_txn = df.groupby(df['Date'].dt.date).size()
        
        freq_stats = pd.DataFrame({
            'Metric': [
                'Average Transactions per Day',
                'Maximum Transactions in a Day',
                'Minimum Transactions in a Day',
                'Days with No Transactions',
                'Most Active Day'
            ],
            'Value': [
                f"{daily_txn.mean():.1f}",
                str(daily_txn.max()),
                str(daily_txn.min()),
                str(30 - len(daily_txn)),
                f"{daily_txn.idxmax()} ({daily_txn.max()} transactions)"
            ]
        })
        
        self.summary_data['transaction_frequency'] = freq_stats
        
        print("\n" + "=" * 70)
        print("TRANSACTION FREQUENCY")
        print("=" * 70)
        print(freq_stats.to_string(index=False))
        print("=" * 70)

    def top_expensive_transactions(self, n=10):
        """
        Display top N most expensive debit transactions.
        Helps identify large purchases for review.
        """
        df = self.df[self.df['Type']=='Debit'].copy()
        top_expensive = df.nlargest(n, 'Amount')[['Date', 'Merchant', 'Amount', 'Time']]
        top_expensive['Date'] = top_expensive['Date'].dt.date
        top_expensive = top_expensive.reset_index(drop=True)
        top_expensive.index = top_expensive.index + 1
        
        self.summary_data['top_expensive_transactions'] = top_expensive
        
        print(f"\n" + "=" * 70)
        print(f"TOP {n} MOST EXPENSIVE TRANSACTIONS")
        print("=" * 70)
        for i, row in top_expensive.iterrows():
            print(f"{i:2d}. ₹{row['Amount']:>10,.2f} | {row['Merchant']:<30} | {row['Date']}")
        print("=" * 70)

    def savings_potential(self):
        """
        Identify potential savings opportunities.
        Analyzes small transactions and frequent merchants for optimization.
        """
        df = self.df[self.df['Type']=='Debit'].copy()
        
        # Small transactions that add up
        small_txn = df[df['Amount'] < 100]
        small_txn_total = small_txn['Amount'].sum()
        small_txn_count = len(small_txn)
        
        # Frequent merchants (potential subscription/regular expenses)
        merchant_freq = df.groupby('Merchant').size().sort_values(ascending=False).head(5)
        
        savings_df = pd.DataFrame({
            'Insight Type': ['Small Transactions', 'Transaction Count', 'Average Amount'],
            'Value': [
                f"₹{small_txn_total:,.2f}",
                str(small_txn_count),
                f"₹{small_txn_total/small_txn_count:.2f}" if small_txn_count > 0 else "N/A"
            ]
        })
        self.summary_data['savings_insights'] = savings_df
        
        print("\n" + "=" * 70)
        print("SAVINGS POTENTIAL INSIGHTS")
        print("=" * 70)
        print(f"\nSmall Transactions (<₹100):")
        print(f"  Total Amount: ₹{small_txn_total:,.2f}")
        print(f"  Count: {small_txn_count} transactions")
        if small_txn_count > 0:
            print(f"  Average: ₹{small_txn_total/small_txn_count:.2f}")
        print(f"\n💡 Tip: Small purchases add up! Consider tracking daily expenses.")
        
        print(f"\nMost Frequent Merchants (Potential Regular Expenses):")
        print("─" * 70)
        for i, (merchant, count) in enumerate(merchant_freq.items(), 1):
            print(f"  {i}. {merchant:<40} {count} transactions")
        print(f"\n💡 Tip: Review frequent merchants for subscription optimization.")
        print("=" * 70)

    def run_all(self):
        """Execute all single month analysis methods and generate visualizations."""
        print("\n" + "🔍 " * 35)
        print("SINGLE MONTH ANALYSIS (LAST 30 DAYS)")
        print("🔍 " * 35 + "\n")
        
        self.summary_stats()
        self.top_merchants()
        self.spending_categories()
        self.weekday_analysis()
        self.time_of_day_analysis()
        self.transaction_frequency()
        self.top_expensive_transactions()
        self.savings_potential()
        
        print("\n" + "📊 " * 35)
        print("GENERATING VISUALIZATIONS")
        print("📊 " * 35 + "\n")
        
        self.plot_daily_spend()
        self.plot_debit_vs_credit()


class MultiMonthAnalysis:
    """
    Analyzes transactions across multiple months.
    
    Provides comprehensive analysis including:
    - Overall summary across all months
    - Month-by-month spending breakdown
    - Spending trends and comparisons
    - Top merchants (overall and per month)
    - Category distribution
    - Savings insights
    - Multiple visualization charts
    """
    
    def __init__(self, df):
        """
        Initialize multi-month analysis.
        
        Args:
            df: DataFrame containing transaction data spanning multiple months
        """
        self.df = df.copy()
        self.df['Month'] = self.df['Date'].dt.strftime('%B %Y')
        self.df['MonthSort'] = self.df['Date'].dt.to_period('M')
        self.summary_data = {}
        self.plots = {}
        
        self.months_sorted = sorted(self.df['MonthSort'].unique())
        self.month_names = [m.strftime('%B %Y') for m in self.months_sorted]

    def overall_summary(self):
        """
        Calculate and display comprehensive overview of all months.
        Includes totals, averages, net flow, and transaction counts.
        """
        df = self.df
        debit_df = df[df['Type']=='Debit']
        credit_df = df[df['Type']=='Credit']
        
        total_debit = debit_df['Amount'].sum()
        total_credit = credit_df['Amount'].sum()
        net_flow = total_credit - total_debit
        
        total_months = len(self.months_sorted)
        avg_monthly_spend = total_debit / total_months if total_months > 0 else 0
        
        total_txn = len(df)
        debit_count = len(debit_df)
        credit_count = len(credit_df)
        
        # Date range
        start_date = df['Date'].min()
        end_date = df['Date'].max()
        days_covered = (end_date - start_date).days + 1
        
        summary_df = pd.DataFrame({
            'Metric': [
                'Analysis Period',
                'Total Months Covered',
                'Total Days Covered',
                '─' * 50,
                'Total Debit (All Time)',
                'Total Credit (All Time)',
                'Net Cash Flow',
                '─' * 50,
                'Average Monthly Spending',
                'Average Daily Spending',
                '─' * 50,
                'Total Transactions',
                'Debit Transactions',
                'Credit Transactions',
                'Average Transactions per Month'
            ],
            'Value': [
                f"{start_date.date()} to {end_date.date()}",
                str(total_months),
                str(days_covered),
                '',
                f"₹{total_debit:,.2f}",
                f"₹{total_credit:,.2f}",
                f"₹{net_flow:,.2f}" + (" (Surplus)" if net_flow > 0 else " (Deficit)" if net_flow < 0 else " (Balanced)"),
                '',
                f"₹{avg_monthly_spend:,.2f}",
                f"₹{total_debit/days_covered:,.2f}",
                '',
                str(total_txn),
                str(debit_count),
                str(credit_count),
                f"{total_txn/total_months:.1f}"
            ]
        })
        
        self.summary_data['overall_summary'] = summary_df
        
        print("=" * 70)
        print("OVERALL SUMMARY - ALL MONTHS")
        print("=" * 70)
        print(f"\nAnalysis Period: {start_date.date()} to {end_date.date()}")
        print(f"Total Months: {total_months} | Total Days: {days_covered}")
        print(f"\n{'─' * 70}")
        print(f"Total Debit:  ₹{total_debit:>15,.2f}")
        print(f"Total Credit: ₹{total_credit:>15,.2f}")
        print(f"Net Flow:     ₹{net_flow:>15,.2f}" + (" (Surplus)" if net_flow > 0 else " (Deficit)" if net_flow < 0 else " (Balanced)"))
        print(f"{'─' * 70}")
        print(f"Average Monthly Spending: ₹{avg_monthly_spend:,.2f}")
        print(f"Average Daily Spending:   ₹{total_debit/days_covered:,.2f}")
        print(f"{'─' * 70}")
        print(f"Total Transactions: {total_txn} (Debit: {debit_count}, Credit: {credit_count})")
        print(f"Avg Transactions/Month: {total_txn/total_months:.1f}")
        print("=" * 70)

    def monthly_spending(self):
        """Detailed month-by-month spending analysis"""
        debit_df = self.df[self.df['Type']=='Debit'].copy()
        
        monthly_stats = []
        for month_period in self.months_sorted:
            month_name = month_period.strftime('%B %Y')
            month_data = debit_df[debit_df['MonthSort'] == month_period]
            
            if len(month_data) > 0:
                total = month_data['Amount'].sum()
                count = len(month_data)
                avg = month_data['Amount'].mean()
                median = month_data['Amount'].median()
                max_amt = month_data['Amount'].max()
                min_amt = month_data['Amount'].min()
                
                monthly_stats.append({
                    'Month': month_name,
                    'Total Spent (INR)': round(total, 2),
                    'Transactions': count,
                    'Average (INR)': round(avg, 2),
                    'Median (INR)': round(median, 2),
                    'Max (INR)': round(max_amt, 2),
                    'Min (INR)': round(min_amt, 2)
                })
        
        monthly_df = pd.DataFrame(monthly_stats)
        self.summary_data['monthly_detailed'] = monthly_df
        
        print("\n" + "=" * 70)
        print("MONTHLY SPENDING BREAKDOWN")
        print("=" * 70)
        print(monthly_df.to_string(index=False))
        print("=" * 70)

    def spending_trends(self):
        """Analyze spending trends across months"""
        debit_df = self.df[self.df['Type']=='Debit'].copy()
        
        monthly_totals = []
        for month_period in self.months_sorted:
            month_name = month_period.strftime('%B %Y')
            month_data = debit_df[debit_df['MonthSort'] == month_period]
            monthly_totals.append(month_data['Amount'].sum())
        
        if len(monthly_totals) > 1:
            # Calculate month-over-month changes
            changes = []
            for i in range(1, len(monthly_totals)):
                change = monthly_totals[i] - monthly_totals[i-1]
                pct_change = (change / monthly_totals[i-1] * 100) if monthly_totals[i-1] > 0 else 0
                changes.append({
                    'From': self.month_names[i-1],
                    'To': self.month_names[i],
                    'Change (INR)': round(change, 2),
                    'Change (%)': round(pct_change, 1),
                    'Trend': '📈 Increase' if change > 0 else '📉 Decrease' if change < 0 else '➡️ Same'
                })
            
            trends_df = pd.DataFrame(changes)
            self.summary_data['spending_trends'] = trends_df
            
            # Find highest and lowest spending months
            max_idx = monthly_totals.index(max(monthly_totals))
            min_idx = monthly_totals.index(min(monthly_totals))
            
            print("\n" + "=" * 70)
            print("SPENDING TRENDS")
            print("=" * 70)
            print(f"\nHighest Spending Month: {self.month_names[max_idx]} (₹{monthly_totals[max_idx]:,.2f})")
            print(f"Lowest Spending Month:  {self.month_names[min_idx]} (₹{monthly_totals[min_idx]:,.2f})")
            print(f"Difference: ₹{monthly_totals[max_idx] - monthly_totals[min_idx]:,.2f}")
            
            print("\n" + "-" * 70)
            print("Month-over-Month Changes:")
            print("-" * 70)
            print(trends_df.to_string(index=False))
            print("=" * 70)

    def top_merchants(self, n=10):
        """Top merchants across all months"""
        top = (self.df[self.df['Type']=='Debit']
               .groupby('Merchant')['Amount']
               .agg(['sum', 'count', 'mean'])
               .sort_values('sum', ascending=False)
               .head(n))
        
        top_df = top.reset_index()
        top_df.columns = ['Merchant', 'Total Spent (INR)', 'Transactions', 'Average (INR)']
        top_df['Total Spent (INR)'] = top_df['Total Spent (INR)'].round(2)
        top_df['Average (INR)'] = top_df['Average (INR)'].round(2)
        
        self.summary_data['top_merchants'] = top_df
        
        print("\n" + "=" * 70)
        print(f"TOP {n} MERCHANTS (ALL TIME)")
        print("=" * 70)
        print(top_df.to_string(index=False))
        print("=" * 70)

    def top_merchants_per_month(self, n=3):
        """Top merchants for each month"""
        all_top_merchants = []
        
        print("\n" + "=" * 70)
        print(f"TOP {n} MERCHANTS PER MONTH")
        print("=" * 70)
        
        for month_period in self.months_sorted:
            month_name = month_period.strftime('%B %Y')
            month_data = self.df[self.df['MonthSort'] == month_period]
            
            top = (month_data[month_data['Type']=='Debit']
                   .groupby('Merchant')['Amount']
                   .sum()
                   .sort_values(ascending=False)
                   .head(n))
            
            print(f"\n{month_name}:")
            print("-" * 70)
            for i, (merchant, amount) in enumerate(top.items(), 1):
                print(f"  {i}. {merchant:<40} ₹{amount:>12,.2f}")
                all_top_merchants.append({
                    'Month': month_name,
                    'Rank': i,
                    'Merchant': merchant,
                    'Amount (INR)': round(amount, 2)
                })
        
        self.summary_data['top_merchants_monthly'] = pd.DataFrame(all_top_merchants)
        print("=" * 70)

    def biggest_transaction_per_month(self):
        """Largest transaction in each month"""
        biggest_txns = []
        
        print("\n" + "=" * 70)
        print("BIGGEST TRANSACTION PER MONTH")
        print("=" * 70)
        
        for month_period in self.months_sorted:
            month_name = month_period.strftime('%B %Y')
            month_data = self.df[self.df['MonthSort'] == month_period]
            
            if len(month_data) > 0:
                max_txn = month_data.loc[month_data['Amount'].idxmax()]
                print(f"\n{month_name}:")
                print(f"  Amount:   ₹{max_txn['Amount']:,.2f}")
                print(f"  Merchant: {max_txn['Merchant']}")
                print(f"  Date:     {max_txn['Date'].date()}")
                print(f"  Type:     {max_txn['Type']}")
                
                biggest_txns.append({
                    'Month': month_name,
                    'Amount (INR)': max_txn['Amount'],
                    'Merchant': max_txn['Merchant'],
                    'Date': max_txn['Date'].date(),
                    'Type': max_txn['Type']
                })
        
        self.summary_data['biggest_transactions'] = pd.DataFrame(biggest_txns)
        print("=" * 70)

    def spending_categories_overall(self):
        """Spending categories across all months"""
        df = self.df[self.df['Type']=='Debit'].copy()
        
        def categorize(amount):
            if amount < 100:
                return 'Under ₹100'
            elif amount < 500:
                return '₹100-500'
            elif amount < 1000:
                return '₹500-1000'
            elif amount < 5000:
                return '₹1000-5000'
            else:
                return 'Above ₹5000'
        
        df['Category'] = df['Amount'].apply(categorize)
        
        category_summary = df.groupby('Category').agg({
            'Amount': ['sum', 'count', 'mean']
        }).round(2)
        category_summary.columns = ['Total (INR)', 'Count', 'Average (INR)']
        category_summary = category_summary.reset_index()
        
        # Sort by category order
        category_order = ['Under ₹100', '₹100-500', '₹500-1000', '₹1000-5000', 'Above ₹5000']
        category_summary['Category'] = pd.Categorical(category_summary['Category'], categories=category_order, ordered=True)
        category_summary = category_summary.sort_values('Category')
        
        self.summary_data['spending_categories'] = category_summary
        
        print("\n" + "=" * 70)
        print("SPENDING BY CATEGORY (ALL TIME)")
        print("=" * 70)
        print(category_summary.to_string(index=False))
        print("=" * 70)

    def monthly_comparison(self):
        """Compare months side by side"""
        debit_df = self.df[self.df['Type']=='Debit'].copy()
        
        comparison = []
        for month_period in self.months_sorted:
            month_name = month_period.strftime('%B %Y')
            month_data = debit_df[debit_df['MonthSort'] == month_period]
            
            if len(month_data) > 0:
                comparison.append({
                    'Month': month_name,
                    'Total Spent': f"₹{month_data['Amount'].sum():,.2f}",
                    'Transactions': len(month_data),
                    'Avg/Transaction': f"₹{month_data['Amount'].mean():,.2f}",
                    'Highest': f"₹{month_data['Amount'].max():,.2f}",
                    'Lowest': f"₹{month_data['Amount'].min():,.2f}"
                })
        
        comparison_df = pd.DataFrame(comparison)
        self.summary_data['monthly_comparison'] = comparison_df
        
        print("\n" + "=" * 70)
        print("MONTH-BY-MONTH COMPARISON")
        print("=" * 70)
        print(comparison_df.to_string(index=False))
        print("=" * 70)

    def savings_insights(self):
        """Savings opportunities across all months"""
        df = self.df[self.df['Type']=='Debit'].copy()
        
        # Small transactions
        small_txn = df[df['Amount'] < 100]
        small_total = small_txn['Amount'].sum()
        small_count = len(small_txn)
        
        # Frequent merchants (potential subscriptions)
        merchant_freq = df.groupby('Merchant').agg({
            'Amount': ['sum', 'count', 'mean']
        }).sort_values(('Amount', 'count'), ascending=False).head(10)
        
        merchant_freq_df = merchant_freq.reset_index()
        merchant_freq_df.columns = ['Merchant', 'Total Spent (INR)', 'Frequency', 'Average (INR)']
        merchant_freq_df = merchant_freq_df.round(2)
        
        self.summary_data['frequent_merchants'] = merchant_freq_df
        
        print("\n" + "=" * 70)
        print("SAVINGS INSIGHTS")
        print("=" * 70)
        print(f"\nSmall Transactions (<₹100):")
        print(f"  Total Amount: ₹{small_total:,.2f}")
        print(f"  Count: {small_count} transactions")
        print(f"  Average: ₹{small_total/small_count:.2f}" if small_count > 0 else "  No small transactions")
        print(f"\n💡 Tip: Small purchases add up! Consider tracking daily expenses.")
        
        print(f"\nMost Frequent Merchants (Potential Regular Expenses):")
        print("-" * 70)
        print(merchant_freq_df.head(5).to_string(index=False))
        print(f"\n💡 Tip: Review frequent merchants for subscription optimization.")
        print("=" * 70)

    def plot_monthly_debit_vs_credit(self):
        """Plot monthly debit vs credit comparison"""
        # Prepare data with sorted months
        monthly_data = []
        for month_period in self.months_sorted:
            month_name = month_period.strftime('%B %Y')
            month_df = self.df[self.df['MonthSort'] == month_period]
            
            debit = month_df[month_df['Type']=='Debit']['Amount'].sum()
            credit = month_df[month_df['Type']=='Credit']['Amount'].sum()
            
            monthly_data.append({
                'Month': month_name,
                'Debit': debit,
                'Credit': credit
            })
        
        plot_df = pd.DataFrame(monthly_data)
        
        fig, ax = plt.subplots(figsize=(12, 6))
        x = range(len(plot_df))
        width = 0.35
        
        ax.bar([i - width/2 for i in x], plot_df['Debit'], width, label='Debit', color='#FF6B6B')
        ax.bar([i + width/2 for i in x], plot_df['Credit'], width, label='Credit', color='#4ECDC4')
        
        ax.set_xlabel('Month')
        ax.set_ylabel('Amount (INR)')
        ax.set_title('Monthly Debit vs Credit Comparison', fontsize=14, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels(plot_df['Month'], rotation=45, ha='right')
        ax.legend()
        ax.grid(axis='y', alpha=0.3)
        
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['monthly_debit_credit'] = buf
        
        plt.savefig('monthly_debit_vs_credit.png')
        print("\n📊 Saved plot: monthly_debit_vs_credit.png")
        plt.show()

    def plot_spending_trend(self):
        """Plot spending trend line across months"""
        monthly_data = []
        for month_period in self.months_sorted:
            month_name = month_period.strftime('%B %Y')
            month_df = self.df[self.df['MonthSort'] == month_period]
            debit = month_df[month_df['Type']=='Debit']['Amount'].sum()
            monthly_data.append({'Month': month_name, 'Spending': debit})
        
        plot_df = pd.DataFrame(monthly_data)
        
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(plot_df['Month'], plot_df['Spending'], marker='o', linewidth=2, markersize=8, color='#FF6B6B')
        ax.fill_between(range(len(plot_df)), plot_df['Spending'], alpha=0.3, color='#FF6B6B')
        
        ax.set_xlabel('Month')
        ax.set_ylabel('Total Spending (INR)')
        ax.set_title('Monthly Spending Trend', fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3)
        plt.xticks(rotation=45, ha='right')
        
        # Add value labels on points
        for i, row in plot_df.iterrows():
            ax.annotate(f'₹{row["Spending"]:,.0f}', 
                       xy=(i, row['Spending']), 
                       xytext=(0, 10), 
                       textcoords='offset points',
                       ha='center',
                       fontsize=9)
        
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['spending_trend'] = buf
        
        plt.savefig('spending_trend.png')
        print("📊 Saved plot: spending_trend.png")
        plt.show()

    def plot_cumulative_spending(self):
        """Plot cumulative spending over time"""
        cumulative = self.df[self.df['Type']=='Debit'].sort_values('Date').copy()
        cumulative['Cumulative'] = cumulative['Amount'].cumsum()
        
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(cumulative['Date'], cumulative['Cumulative'], linewidth=2, color='#4ECDC4')
        ax.fill_between(cumulative['Date'], cumulative['Cumulative'], alpha=0.3, color='#4ECDC4')
        
        ax.set_title("Cumulative Spending Over Time", fontsize=14, fontweight='bold')
        ax.set_ylabel("Cumulative Amount (INR)")
        ax.set_xlabel("Date")
        ax.grid(True, alpha=0.3)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['cumulative_spending'] = buf
        
        plt.savefig('cumulative_spending.png')
        print("📊 Saved plot: cumulative_spending.png")
        plt.show()

    def plot_debit_credit_ratio(self):
        """Plot overall debit vs credit ratio"""
        type_summary = self.df.groupby('Type')['Amount'].sum()
        
        fig, ax = plt.subplots(figsize=(8, 8))
        colors = ['#FF6B6B', '#4ECDC4']
        explode = (0.05, 0)
        
        ax.pie(type_summary, labels=type_summary.index, autopct='%1.1f%%', 
               colors=colors, explode=explode, shadow=True, startangle=90)
        ax.set_title('Overall Debit vs Credit Ratio', fontsize=14, fontweight='bold')
        
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['debit_credit_ratio'] = buf
        
        plt.savefig('debit_credit_ratio.png')
        print("📊 Saved plot: debit_credit_ratio.png")
        plt.show()

    def plot_category_distribution(self):
        """Plot spending distribution by category"""
        df = self.df[self.df['Type']=='Debit'].copy()
        
        def categorize(amount):
            if amount < 100:
                return 'Under ₹100'
            elif amount < 500:
                return '₹100-500'
            elif amount < 1000:
                return '₹500-1000'
            elif amount < 5000:
                return '₹1000-5000'
            else:
                return 'Above ₹5000'
        
        df['Category'] = df['Amount'].apply(categorize)
        category_totals = df.groupby('Category')['Amount'].sum()
        
        # Sort by category order
        category_order = ['Under ₹100', '₹100-500', '₹500-1000', '₹1000-5000', 'Above ₹5000']
        category_totals = category_totals.reindex(category_order)
        
        fig, ax = plt.subplots(figsize=(10, 6))
        colors = ['#FFD93D', '#6BCB77', '#4D96FF', '#FF6B9D', '#C44569']
        bars = ax.bar(category_totals.index, category_totals.values, color=colors)
        
        ax.set_xlabel('Spending Category')
        ax.set_ylabel('Total Amount (INR)')
        ax.set_title('Spending Distribution by Category', fontsize=14, fontweight='bold')
        ax.grid(axis='y', alpha=0.3)
        plt.xticks(rotation=45, ha='right')
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'₹{height:,.0f}',
                   ha='center', va='bottom', fontsize=9)
        
        plt.tight_layout()
        
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        self.plots['category_distribution'] = buf
        
        plt.savefig('category_distribution.png')
        print("📊 Saved plot: category_distribution.png")
        plt.show()

    def run_all(self):
        print("\n" + "🔍 " * 35)
        print("MULTI-MONTH ANALYSIS")
        print("🔍 " * 35 + "\n")
        
        self.overall_summary()
        self.monthly_spending()
        self.spending_trends()
        self.monthly_comparison()
        self.top_merchants()
        self.top_merchants_per_month()
        self.biggest_transaction_per_month()
        self.spending_categories_overall()
        self.savings_insights()
        
        print("\n" + "📊 " * 35)
        print("GENERATING VISUALIZATIONS")
        print("📊 " * 35 + "\n")
        
        self.plot_monthly_debit_vs_credit()
        self.plot_spending_trend()
        self.plot_cumulative_spending()
        self.plot_category_distribution()
        self.plot_debit_credit_ratio()


def load_pdf(pdf_path, pdf_password):
    """
    Load and extract text from a PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        pdf_password: Password for encrypted PDFs (None if not password-protected)
    
    Returns:
        str: Extracted text from all pages of the PDF
    
    Raises:
        SystemExit: If PDF cannot be opened or password is incorrect
    """
    try:
        with pdfplumber.open(pdf_path, password=pdf_password) as pdf:
            text = "\n".join(page.extract_text() for page in pdf.pages)
        print("PDF opened successfully!")
        return text
    except pdfplumber.pdf.PDFPasswordIncorrect:
        print("Incorrect PDF password. Please check and try again.")
        sys.exit(1)
    except Exception as e:
        print(f"Failed to open PDF: {e}")
        sys.exit(1)


def detect_pdf_type(text):
    """
    Automatically detect whether PDF is from PhonePe or Google Pay.
    
    Args:
        text: Extracted text from PDF
    
    Returns:
        str: "PhonePe" or "GooglePay"
    
    Raises:
        ValueError: If PDF format cannot be detected
    """
    clean_text = " ".join(text.split())
    
    if "Transaction ID :" in clean_text and "Debited from" in clean_text:
        pdf_type = "PhonePe"
    elif "Paidto" in clean_text and "UPITransactionID:" in clean_text:
        pdf_type = "GooglePay"
    else:
        raise ValueError("Unsupported PDF format! Could not detect PhonePe or Google Pay format.")
    
    print(f"Detected PDF type: {pdf_type}")
    return pdf_type


def add_spaces_to_name(name):
    """
    Add spaces to concatenated merchant names from Google Pay PDFs.
    
    Google Pay PDFs often have names without spaces (e.g., 'MissRUCHIKASUBHASHPANDE').
    This function intelligently adds spaces between words based on capitalization patterns.
    
    Args:
        name: Merchant name string (potentially without spaces)
    
    Returns:
        str: Name with proper spacing (e.g., 'Miss RUCHIKA SUBHASH PANDE')
    
    Example:
        'MissRUCHIKASUBHASHPANDE' -> 'Miss RUCHIKA SUBHASH PANDE'
        'StateBankofIndia' -> 'State Bankof India'
    """
    if not name or len(name) < 3:
        return name
    
    if ' ' in name:
        return name
    
    result = []
    for i, char in enumerate(name):
        if i > 0 and char.isupper() and name[i-1].islower():
            result.append(' ')
        result.append(char)
    
    spaced_name = ''.join(result)
    
    spaced_name = spaced_name.replace('Mr ', 'Mr. ')
    spaced_name = spaced_name.replace('Mrs ', 'Mrs. ')
    spaced_name = spaced_name.replace('Miss ', 'Miss ')
    spaced_name = spaced_name.replace('Dr ', 'Dr. ')
    
    return spaced_name


def parse_transactions(text, pdf_type):
    """
    Parse transactions from PDF text using regex patterns.
    
    Extracts transaction details including date, merchant, amount, type, etc.
    Uses different regex patterns for PhonePe and Google Pay formats.
    
    Args:
        text: Extracted text from PDF
        pdf_type: "PhonePe" or "GooglePay"
    
    Returns:
        list: List of transaction dictionaries with keys:
              Date, Time, Merchant, Type, Amount, Transaction_ID, Account
    """
    transactions = []
    
    if pdf_type == "PhonePe":
        pattern = re.compile(
            r"([A-Za-z]{3}\s\d{2},\s\d{4})\s+"
            r"(?:Paid to|Received from)\s+(.*?)\s+"
            r"(Debit|Credit)\s+INR\s+([\d,]+\.\d{2})\s+"
            r"([\d:APM\s]+)\s+"
            r"Transaction ID : ([A-Z0-9]+)\s+"
            r"UTR No : (\d+)\s+"
            r"(?:Debited from|Credited to)\s+(XX\d+)",
            re.DOTALL
        )
        for match in re.finditer(pattern, text):
            date, merchant, txn_type, amount, time, txn_id, utr, account = match.groups()
            transactions.append({
                "Date": pd.to_datetime(date),
                "Time": time.strip(),
                "Merchant": merchant.strip(),
                "Type": txn_type,
                "Amount": float(amount.replace(",", "")),
                "Transaction_ID": txn_id,
                "Account": account.strip()
            })
    
    elif pdf_type == "GooglePay":
        text = re.sub(r'(?<!\s)(Paidto)', r' Paidto', text)
        text = re.sub(r'(?<!\s)(UPITransactionID:)', r' UPITransactionID:', text)
        text = re.sub(r'(?<!\s)(Paidby)', r' Paidby', text)
        text = re.sub(r'(?<!\s)(Receivedfrom)', r' Receivedfrom', text)
        
        pattern = re.compile(
            r"(\d{2}[A-Za-z]{3},\d{4})\s*"                    # Date
            r"(Paidto|Receivedfrom)\s*"                       # Transaction type
            r"(.*?)\s*"                                        # Merchant (non-greedy)
            r"₹([\d,]+\.?\d*)\s*"                             # Amount (flexible decimal)
            r"([\d:APM]+)?\s*"                                 # Time (optional)
            r"UPI\s*Transaction\s*ID:?\s*([\d]+)",            # Transaction ID
            re.DOTALL
        )
        
        for match in re.finditer(pattern, text):
            date, tx_type_word, merchant, amount, time, txn_id = match.groups()
            tx_type = "Debit" if tx_type_word.strip() == "Paidto" else "Credit"
            
            merchant_clean = merchant.strip()
            
            merchant_clean = re.sub(r'₹[\d,]+\.?\d*', '', merchant_clean)
            merchant_clean = re.sub(r'\d{1,2}:\d{2}[AP]M', '', merchant_clean)
            merchant_clean = re.sub(r'UPI\s*Transaction\s*ID:?\s*[\d]+', '', merchant_clean, flags=re.IGNORECASE)
            merchant_clean = re.sub(r'\b(Paidto|Receivedfrom|Paidby|UPI|Transaction|ID)\b.*', '', merchant_clean, flags=re.IGNORECASE)
            merchant_clean = merchant_clean.split('\n')[0].strip()
            
            if len(merchant_clean) > 50:
                parts = re.split(r'\s{2,}|\t', merchant_clean)
                merchant_clean = parts[0].strip() if parts else merchant_clean[:50]
            
            merchant_clean = ' '.join(merchant_clean.split())
            
            if not merchant_clean:
                merchant_clean = "Unknown"
            
            merchant_clean = add_spaces_to_name(merchant_clean)
            
            transactions.append({
                "Date": pd.to_datetime(date, format='%d%b,%Y'),
                "Time": time.strip() if time else "",
                "Merchant": merchant_clean,
                "Type": tx_type,
                "Amount": float(amount.replace(",", "")),
                "Transaction_ID": txn_id.strip(),
                "Account": ""
            })
    
    return transactions


def save_to_excel(df, analysis):
    """
    Save all analysis results and plots to a single Excel file.
    
    Creates a comprehensive Excel workbook containing:
    - All transactions
    - Summary statistics tables
    - Analysis results
    - Embedded visualization charts
    
    Args:
        df: DataFrame containing all transactions
        analysis: Analysis object (SingleMonthAnalysis or MultiMonthAnalysis)
                 containing summary_data and plots dictionaries
    
    Output:
        Excel file named 'statement_analysis_YYYYMMDD_HHMMSS.xlsx'
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'statement_analysis_{timestamp}.xlsx'
    
    print(f"\n\n{'=' * 70}")
    print(f"SAVING ANALYSIS TO EXCEL")
    print(f"{'=' * 70}")
    print(f"📄 Filename: {output_file}")
    
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            df.to_excel(writer, sheet_name='All Transactions', index=False)
            
            used_names = {'All Transactions'}
            
            for sheet_name, data in analysis.summary_data.items():
                clean_name = sheet_name.replace('_', ' ').title()
                
                original_name = clean_name
                counter = 1
                while clean_name in used_names or clean_name.lower() in [n.lower() for n in used_names]:
                    clean_name = f"{original_name} {counter}"
                    counter += 1
                
                if len(clean_name) > 31:
                    clean_name = clean_name[:31]
                
                used_names.add(clean_name)
                data.to_excel(writer, sheet_name=clean_name, index=False)
            
            for plot_name, plot_buf in analysis.plots.items():
                clean_name = plot_name.replace('_', ' ').title() + ' Chart'
                
                original_name = clean_name
                counter = 1
                while clean_name in used_names or clean_name.lower() in [n.lower() for n in used_names]:
                    clean_name = f"{original_name} {counter}"
                    counter += 1
                
                if len(clean_name) > 31:
                    clean_name = clean_name[:28] + ' Ch'
                
                used_names.add(clean_name)
                worksheet = workbook.add_worksheet(clean_name)
                plot_buf.seek(0)
                worksheet.insert_image('B2', plot_name, {'image_data': plot_buf})
            
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                worksheet.set_column('A:Z', 20)
        
        print(f"\n✅ Successfully saved analysis!")
        print(f"{'─' * 70}")
        print(f"📊 File contains:")
        print(f"   • All transactions")
        print(f"   • Summary statistics")
        print(f"   • Analysis tables")
        print(f"   • All visualization charts")
        print(f"{'─' * 70}")
        print(f"📁 Location: {os.path.abspath(output_file)}")
        print(f"{'=' * 70}")
        
    except Exception as e:
        print(f"❌ Error saving to Excel: {e}")
        print("Note: Make sure 'xlsxwriter' is installed: pip install xlsxwriter")


def main():
    """
    Main entry point for the Statement Analyser.
    
    Workflow:
    1. Accept PDF file path (command-line arg or interactive input)
    2. Load and extract text from PDF (with password support)
    3. Detect PDF type (PhonePe or Google Pay)
    4. Parse transactions using appropriate regex patterns
    5. Determine analysis type (single month vs multi-month)
    6. Run comprehensive analysis and generate visualizations
    7. Export everything to timestamped Excel file
    """
    print("=" * 60)
    print("PhonePe + Google Pay Statement Analyser")
    print("=" * 60)
    
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = input("Enter the path to your PDF file: ").strip()
    
    if not os.path.exists(pdf_path):
        print(f"Error: File '{pdf_path}' not found!")
        sys.exit(1)
    
    pdf_password = getpass.getpass("Enter PDF password (press Enter if no password): ")
    if not pdf_password:
        pdf_password = None
    
    print("\nLoading PDF...")
    text = load_pdf(pdf_path, pdf_password)
    
    print("\nDetecting PDF type...")
    pdf_type = detect_pdf_type(text)
    
    print("\nParsing transactions...")
    transactions = parse_transactions(text, pdf_type)
    
    if not transactions:
        print("No transactions found in the PDF!")
        sys.exit(1)
    
    df = pd.DataFrame(transactions)
    print(f"\nFound {len(df)} transactions")
    print("\nFirst few transactions:")
    print(df.head())
    
    min_date = df['Date'].min()
    max_date = df['Date'].max()
    date_range_days = (max_date - min_date).days + 1
    
    print(f"\nDate range: {min_date.date()} to {max_date.date()} ({date_range_days} days)")
    
    print("\n" + "=" * 60)
    print("ANALYSIS RESULTS")
    print("=" * 60 + "\n")
    
    if date_range_days <= 30:
        print("Running Single Month Analysis...")
        analysis = SingleMonthAnalysis(df)
    else:
        print("Running Multi-Month Analysis...")
        analysis = MultiMonthAnalysis(df)
    
    analysis.run_all()
    
    # Save to Excel
    save_to_excel(df, analysis)
    
    print("\n" + "=" * 60)
    print("Analysis complete!")
    print("=" * 60)


if __name__ == "__main__":
    main()

