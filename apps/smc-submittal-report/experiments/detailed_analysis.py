import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# Read the CSV file
df = pd.read_csv('/mnt/user-data/uploads/SubmittalLog.csv', encoding='utf-8-sig')

# Parse dates properly
date_columns = ['Created At', 'Updated At', 'Sent Date', 'Returned Date', 'Due Date', 
                'Final Due Date', 'Distributed Date']
for col in date_columns:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], format='%d/%m/%Y', errors='coerce')

# Filter for NSW Ports related submittals
nsw_ports_mask = (
    df['Approvers'].str.contains('NSW Ports', na=False, case=False) |
    df['Action Required By'].str.contains('NSW Ports', na=False, case=False)
)
nsw_ports_df = df[nsw_ports_mask].copy()

# Filter for Hold Points
hold_points_df = nsw_ports_df[nsw_ports_df['Type'].str.contains('Hold Point', na=False, case=False)].copy()

# Get week boundaries
today = datetime.now()
days_since_sunday = (today.weekday() + 1) % 7
last_sunday = today - timedelta(days=days_since_sunday)
last_sunday = last_sunday.replace(hour=23, minute=59, second=59)
week_start = last_sunday - timedelta(days=7)
week_start = week_start.replace(hour=0, minute=0, second=0)

print("="*80)
print("NSW PORTS HOLD POINTS - DETAILED WEEKLY ANALYSIS")
print(f"Week: {week_start.strftime('%d %B %Y')} to {last_sunday.strftime('%d %B %Y')}")
print("="*80)
print()

# Contractor Analysis
print("CONTRACTOR BREAKDOWN")
print("-"*40)
contractor_counts = hold_points_df['Responsible Contractor'].value_counts()
for contractor, count in contractor_counts.items():
    print(f"{contractor}: {count} Hold Points")
print()

# Specification Section Analysis
print("SPECIFICATION SECTIONS")
print("-"*40)
spec_counts = hold_points_df['Spec Section'].value_counts()
for spec, count in spec_counts.items():
    print(f"{spec}: {count} Hold Points")
print()

# Pending Items Analysis
pending_df = hold_points_df[hold_points_df['Response'] == 'Pending'].copy()
print("HOLD POINTS AWAITING NSW PORTS ACTION")
print("-"*40)
print(f"Total Pending: {len(pending_df)}")
if len(pending_df) > 0:
    pending_df = pending_df.sort_values('Due Date')
    print("\nUpcoming Due Dates:")
    for idx, row in pending_df.iterrows():
        due_date = row['Due Date'].strftime('%d/%m/%Y') if pd.notna(row['Due Date']) else 'No due date'
        days_until = (row['Due Date'] - datetime.now()).days if pd.notna(row['Due Date']) else 999
        urgency = "OVERDUE" if days_until < 0 else f"{days_until} days" if days_until != 999 else ""
        print(f"  {row['#']}: {due_date} ({urgency}) - {row['Title'][:40]}...")
print()

# Items Not Released Analysis
not_released = hold_points_df[hold_points_df['Response'].str.contains('Not Released', na=False)].copy()
print("HOLD POINTS NOT RELEASED (Requiring Resubmission)")
print("-"*40)
print(f"Total Not Released: {len(not_released)}")
if len(not_released) > 0:
    for idx, row in not_released.iterrows():
        print(f"  {row['#']}: {row['Title'][:50]}...")
        if pd.notna(row['Returned Date']):
            print(f"    Returned: {row['Returned Date'].strftime('%d/%m/%Y')}")
print()

# Released with Conditions Analysis
conditions = hold_points_df[hold_points_df['Response'].str.contains('Released with Conditions', na=False)].copy()
print("HOLD POINTS RELEASED WITH CONDITIONS")
print("-"*40)
print(f"Total Released with Conditions: {len(conditions)}")
if len(conditions) > 0:
    for idx, row in conditions.iterrows():
        print(f"  {row['#']}: {row['Title'][:50]}...")
print()

# Performance Metrics
print("PERFORMANCE METRICS")
print("-"*40)

# Calculate response times for returned items
returned_items = hold_points_df[pd.notna(hold_points_df['Returned Date']) & pd.notna(hold_points_df['Sent Date'])].copy()
if len(returned_items) > 0:
    returned_items['Response Time'] = (returned_items['Returned Date'] - returned_items['Sent Date']).dt.days
    avg_response = returned_items['Response Time'].mean()
    print(f"Average Response Time: {avg_response:.1f} days")
    print(f"Fastest Response: {returned_items['Response Time'].min()} days")
    print(f"Slowest Response: {returned_items['Response Time'].max()} days")
print()

# Weekly Trend
print("SUBMISSION TRENDS (Last 4 Weeks)")
print("-"*40)
for week_num in range(4):
    week_end = last_sunday - timedelta(days=7*week_num)
    week_begin = week_end - timedelta(days=7)
    week_submissions = hold_points_df[
        (hold_points_df['Sent Date'] > week_begin) & 
        (hold_points_df['Sent Date'] <= week_end)
    ]
    week_releases = hold_points_df[
        (hold_points_df['Returned Date'] > week_begin) & 
        (hold_points_df['Returned Date'] <= week_end) &
        (hold_points_df['Response'].str.contains('Released', na=False))
    ]
    print(f"Week ending {week_end.strftime('%d/%m')}: {len(week_submissions)} submitted, {len(week_releases)} released")
print()

# Critical Items
print("CRITICAL ITEMS REQUIRING ATTENTION")
print("-"*40)

# Overdue items
overdue = pending_df[pending_df['Due Date'] < datetime.now()]
if len(overdue) > 0:
    print(f"⚠️  {len(overdue)} Hold Points are OVERDUE for response")
    for idx, row in overdue.iterrows():
        days_overdue = (datetime.now() - row['Due Date']).days
        print(f"   {row['#']}: {days_overdue} days overdue")

# Items due within 2 days
urgent = pending_df[(pending_df['Due Date'] - datetime.now()).dt.days.between(0, 2)]
if len(urgent) > 0:
    print(f"⚠️  {len(urgent)} Hold Points due within 2 days")
    for idx, row in urgent.iterrows():
        print(f"   {row['#']}: Due {row['Due Date'].strftime('%d/%m/%Y')}")

# Multiple resubmissions
resubmitted = hold_points_df.groupby('#').size()
multiple_submissions = resubmitted[resubmitted > 1]
if len(multiple_submissions) > 0:
    print(f"⚠️  {len(multiple_submissions)} Hold Points have been submitted multiple times")
    for hp_num, count in multiple_submissions.items():
        print(f"   {hp_num}: {count} submissions")

print()
print("="*80)
print("END OF DETAILED ANALYSIS")
print("="*80)
