import pandas as pd
import os
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Define the base directory
base_dir = Path('/Users/bisratgizaw/Downloads/summary dashboard')

# Months to analyze
months = ['Jan', 'February', 'March', 'April', 'May', 'June', 'July', 'August']

# Initialize storage for analysis
all_data = []
monthly_stats = {}
campaign_types = {}
program_data = {}

print("=" * 80)
print("CALL CENTER DATA ANALYSIS - JANUARY TO AUGUST 2025")
print("=" * 80)
print()

# Function to extract info from filename
def extract_campaign_info(filename):
    """Extract campaign type and program from filename"""
    filename_lower = filename.lower()
    
    # Campaign types
    if 'payment' in filename_lower or 'pay' in filename_lower:
        campaign_type = 'Payment Nudge'
    elif 'application' in filename_lower or 'applicant' in filename_lower or 'nudge' in filename_lower:
        campaign_type = 'Application Nudge'
    elif 'assignment' in filename_lower or 'submission' in filename_lower:
        campaign_type = 'Assignment Follow-up'
    elif 'event' in filename_lower or 'info session' in filename_lower or 'rsvp' in filename_lower or 'invitation' in filename_lower or 'invite' in filename_lower:
        campaign_type = 'Event Invitation'
    elif 'inprogress' in filename_lower or 'in-progress' in filename_lower:
        campaign_type = 'In-Progress Follow-up'
    elif 'hackathon' in filename_lower or 'hachaton' in filename_lower:
        campaign_type = 'Hackathon'
    elif 'enrolled' in filename_lower:
        campaign_type = 'Enrollment Follow-up'
    elif 'retry' in filename_lower or 'not reached' in filename_lower:
        campaign_type = 'Retry/Not Reached'
    else:
        campaign_type = 'Other'
    
    # Program identification
    programs = []
    if 'aice' in filename_lower or 'ai career' in filename_lower:
        programs.append('AiCE')
    if 'pathway' in filename_lower or 'pw' in filename_lower or 'pf' in filename_lower and 'professional foundation' not in filename_lower:
        programs.append('Pathway')
    if 'professional foundation' in filename_lower:
        programs.append('Professional Foundation')
    if 'virtual assistant' in filename_lower or 'va' in filename_lower:
        programs.append('Virtual Assistant')
    if 'aws' in filename_lower or 'cloud' in filename_lower:
        programs.append('AWS')
    if 'data analytics' in filename_lower or 'da ' in filename_lower:
        programs.append('Data Analytics')
    if 'data science' in filename_lower or 'ds ' in filename_lower:
        programs.append('Data Science')
    if 'front-end' in filename_lower or 'fe ' in filename_lower or 'frontend' in filename_lower:
        programs.append('Front-End')
    if 'back-end' in filename_lower or 'be ' in filename_lower or 'backend' in filename_lower:
        programs.append('Back-End')
    if 'software' in filename_lower or 'sf ' in filename_lower:
        programs.append('Software Engineering')
    if 'techlite' in filename_lower or 'tech-lite' in filename_lower:
        programs.append('Techlite')
    if 'incubation' in filename_lower or 'karibu' in filename_lower:
        programs.append('Incubation')
    if 'content creation' in filename_lower or 'cc ' in filename_lower:
        programs.append('Content Creation')
    if 'alu' in filename_lower:
        programs.append('ALU Partnership')
    
    if not programs:
        programs = ['General']
    
    return campaign_type, programs

# Process each month
for month in months:
    month_dir = base_dir / month
    
    if not month_dir.exists():
        print(f"⚠️  {month} folder not found, skipping...")
        continue
    
    print(f"\n{'=' * 80}")
    print(f"ANALYZING {month.upper()} 2025")
    print(f"{'=' * 80}")
    
    # Get all Excel files
    excel_files = list(month_dir.glob('*.xlsx')) + list(month_dir.glob('*.xls'))
    excel_files = [f for f in excel_files if not f.name.startswith('~$')]  # Exclude temp files
    
    print(f"Found {len(excel_files)} campaign files")
    
    month_contacts = 0
    month_campaigns = 0
    month_campaign_types = {}
    month_programs = {}
    
    for file in excel_files:
        try:
            # Read Excel file
            df = pd.read_excel(file)
            
            # Count rows (excluding header)
            contacts = len(df)
            
            # Extract campaign info
            campaign_type, programs = extract_campaign_info(file.name)
            
            # Track data
            month_contacts += contacts
            month_campaigns += 1
            
            # Campaign type tracking
            month_campaign_types[campaign_type] = month_campaign_types.get(campaign_type, 0) + contacts
            campaign_types[campaign_type] = campaign_types.get(campaign_type, 0) + contacts
            
            # Program tracking
            for prog in programs:
                month_programs[prog] = month_programs.get(prog, 0) + contacts
                program_data[prog] = program_data.get(prog, 0) + contacts
            
            # Store file info
            all_data.append({
                'Month': month,
                'Campaign': file.name,
                'Type': campaign_type,
                'Programs': ', '.join(programs),
                'Contacts': contacts
            })
            
        except Exception as e:
            print(f"  ⚠️  Error reading {file.name}: {str(e)}")
    
    # Store monthly stats
    monthly_stats[month] = {
        'total_contacts': month_contacts,
        'total_campaigns': month_campaigns,
        'campaign_types': month_campaign_types,
        'programs': month_programs
    }
    
    # Print monthly summary
    print(f"\n📊 {month} Summary:")
    print(f"  Total Contacts: {month_contacts:,}")
    print(f"  Total Campaigns: {month_campaigns}")
    print(f"\n  Top Campaign Types:")
    for ctype, count in sorted(month_campaign_types.items(), key=lambda x: x[1], reverse=True)[:5]:
        print(f"    - {ctype}: {count:,}")
    print(f"\n  Top Programs:")
    for prog, count in sorted(month_programs.items(), key=lambda x: x[1], reverse=True)[:5]:
        print(f"    - {prog}: {count:,}")

# Generate comprehensive report
print("\n\n")
print("=" * 80)
print("COMPREHENSIVE SUMMARY REPORT - JANUARY TO AUGUST 2025")
print("=" * 80)

# Overall Statistics
total_contacts = sum(stats['total_contacts'] for stats in monthly_stats.values())
total_campaigns = sum(stats['total_campaigns'] for stats in monthly_stats.values())

print(f"\n📈 OVERALL STATISTICS")
print(f"{'─' * 80}")
print(f"Total Period: January - August 2025 (8 months)")
print(f"Total Contacts Made: {total_contacts:,}")
print(f"Total Campaigns Run: {total_campaigns}")
print(f"Average Contacts per Month: {total_contacts//8:,}")
print(f"Average Campaigns per Month: {total_campaigns//8}")
print(f"Average Contacts per Campaign: {total_contacts//total_campaigns if total_campaigns > 0 else 0}")

# Monthly Breakdown
print(f"\n📅 MONTHLY BREAKDOWN")
print(f"{'─' * 80}")
print(f"{'Month':<15} {'Contacts':>12} {'Campaigns':>12} {'Avg/Campaign':>15}")
print(f"{'─' * 80}")
for month in months:
    if month in monthly_stats:
        stats = monthly_stats[month]
        avg = stats['total_contacts'] // stats['total_campaigns'] if stats['total_campaigns'] > 0 else 0
        print(f"{month:<15} {stats['total_contacts']:>12,} {stats['total_campaigns']:>12} {avg:>15}")

# Campaign Type Analysis
print(f"\n🎯 CAMPAIGN TYPE ANALYSIS")
print(f"{'─' * 80}")
print(f"{'Campaign Type':<30} {'Total Contacts':>15} {'% of Total':>12} {'# Campaigns':>12}")
print(f"{'─' * 80}")
for ctype, count in sorted(campaign_types.items(), key=lambda x: x[1], reverse=True):
    campaigns_count = sum(1 for d in all_data if d['Type'] == ctype)
    percentage = (count / total_contacts * 100) if total_contacts > 0 else 0
    print(f"{ctype:<30} {count:>15,} {percentage:>11.1f}% {campaigns_count:>12}")

# Program Analysis
print(f"\n🎓 PROGRAM ANALYSIS")
print(f"{'─' * 80}")
print(f"{'Program':<30} {'Total Contacts':>15} {'% of Total':>12}")
print(f"{'─' * 80}")
for prog, count in sorted(program_data.items(), key=lambda x: x[1], reverse=True):
    percentage = (count / total_contacts * 100) if total_contacts > 0 else 0
    print(f"{prog:<30} {count:>15,} {percentage:>11.1f}%")

# Monthly Trends
print(f"\n📊 MONTHLY TRENDS")
print(f"{'─' * 80}")
month_values = [(month, monthly_stats[month]['total_contacts']) for month in months if month in monthly_stats]
if len(month_values) > 1:
    for i in range(1, len(month_values)):
        curr_month, curr_val = month_values[i]
        prev_month, prev_val = month_values[i-1]
        change = curr_val - prev_val
        change_pct = (change / prev_val * 100) if prev_val > 0 else 0
        direction = "📈" if change > 0 else "📉" if change < 0 else "➡️"
        print(f"{prev_month} → {curr_month}: {direction} {change:+,} ({change_pct:+.1f}%)")

# Key Insights
print(f"\n💡 KEY INSIGHTS")
print(f"{'─' * 80}")

# Busiest month
busiest_month = max(monthly_stats.items(), key=lambda x: x[1]['total_contacts'])
print(f"1. Busiest Month: {busiest_month[0]} with {busiest_month[1]['total_contacts']:,} contacts")

# Most common campaign type
top_campaign = max(campaign_types.items(), key=lambda x: x[1])
print(f"2. Most Common Campaign Type: {top_campaign[0]} ({top_campaign[1]:,} contacts, {top_campaign[1]/total_contacts*100:.1f}%)")

# Top program
top_program = max(program_data.items(), key=lambda x: x[1])
print(f"3. Top Program: {top_program[0]} ({top_program[1]:,} contacts, {top_program[1]/total_contacts*100:.1f}%)")

# Growth analysis
first_month_val = month_values[0][1]
last_month_val = month_values[-1][1]
total_growth = ((last_month_val - first_month_val) / first_month_val * 100) if first_month_val > 0 else 0
print(f"4. Overall Growth: {total_growth:+.1f}% from {month_values[0][0]} to {month_values[-1][0]}")

# Campaign diversity
print(f"5. Campaign Diversity: {len(campaign_types)} different campaign types executed")
print(f"6. Program Coverage: {len(program_data)} programs reached")

# Save detailed data to CSV
df_all = pd.DataFrame(all_data)
output_file = base_dir / 'detailed_analysis_jan_to_aug.csv'
df_all.to_csv(output_file, index=False)

print(f"\n✅ Detailed data saved to: {output_file}")
print(f"\n{'=' * 80}")
print(f"Report generated on: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
print(f"{'=' * 80}")
