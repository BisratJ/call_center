import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Define the base directory
base_dir = Path('/Users/bisratgizaw/Downloads/summary dashboard')
months = ['Jan', 'February', 'March', 'April', 'May', 'June', 'July', 'August']

print("=" * 100)
print("COMPREHENSIVE CALL CENTER ANALYSIS - JANUARY TO AUGUST 2025")
print("Based on Agents Assigned Data with Call Outcomes")
print("=" * 100)
print()

# Storage for analysis
all_contacts_data = []
monthly_detailed_stats = {}
overall_unique_contacts = set()
overall_call_outcomes = {}
overall_interest_levels = {}
overall_delay_reasons = {}

# Function to normalize call status
def normalize_call_status(status):
    if pd.isna(status):
        return 'Unknown'
    status_str = str(status).lower().strip()
    
    if 'reach' in status_str or 'spoke' in status_str or 'contacted' in status_str or 'answered' in status_str:
        return 'Reached'
    elif 'busy' in status_str or 'line busy' in status_str:
        return 'Busy'
    elif 'not pick' in status_str or "didn't pick" in status_str or 'no answer' in status_str or 'no response' in status_str:
        return 'Not Picking Up'
    elif 'not working' in status_str or 'number not' in status_str or 'invalid' in status_str or 'wrong' in status_str or 'switched off' in status_str:
        return 'Phone Not Working'
    elif 'retry' in status_str or 'try again' in status_str:
        return 'Retry'
    else:
        return 'Other'

# Function to normalize interest level
def normalize_interest(interest):
    if pd.isna(interest):
        return 'Unknown'
    interest_str = str(interest).lower().strip()
    
    if 'yes' in interest_str or 'interested' in interest_str or 'continue' in interest_str:
        return 'Interested'
    elif 'considering' in interest_str or 'thinking' in interest_str or 'maybe' in interest_str:
        return 'Considering'
    elif 'no' in interest_str or 'not interested' in interest_str:
        return 'Not Interested'
    else:
        return 'Unknown'

# Function to normalize delay reasons
def normalize_delay_reason(reason):
    if pd.isna(reason):
        return 'Not Specified'
    reason_str = str(reason).lower().strip()
    
    if 'personal' in reason_str or 'family' in reason_str or 'circumstances' in reason_str:
        return 'Personal Circumstances'
    elif 'busy' in reason_str or 'schedule' in reason_str or 'time' in reason_str:
        return 'Busy Schedule'
    elif 'forgot' in reason_str or 'forget' in reason_str:
        return 'Forgot'
    elif 'financial' in reason_str or 'money' in reason_str or 'payment' in reason_str or 'cost' in reason_str:
        return 'Financial Reasons'
    elif 'job' in reason_str or 'work' in reason_str or 'employment' in reason_str:
        return 'Work/Job Related'
    elif 'travel' in reason_str or 'relocat' in reason_str:
        return 'Travel/Relocation'
    elif 'health' in reason_str:
        return 'Health Issues'
    else:
        return 'Other'

# Extract campaign info from filename
def extract_campaign_info(filename):
    filename_lower = filename.lower()
    
    if 'payment' in filename_lower or 'pay' in filename_lower:
        campaign_type = 'Payment Nudge'
    elif 'application' in filename_lower or 'applicant' in filename_lower or 'nudge' in filename_lower:
        campaign_type = 'Application Nudge'
    elif 'event' in filename_lower or 'info session' in filename_lower or 'rsvp' in filename_lower:
        campaign_type = 'Event Invitation'
    elif 'inprogress' in filename_lower or 'in-progress' in filename_lower:
        campaign_type = 'In-Progress Follow-up'
    elif 'enrolled' in filename_lower:
        campaign_type = 'Enrollment Follow-up'
    else:
        campaign_type = 'Other'
    
    programs = []
    if 'pathway' in filename_lower or 'pw' in filename_lower:
        programs.append('Pathway')
    if 'aice' in filename_lower:
        programs.append('AiCE')
    if 'techlite' in filename_lower:
        programs.append('Techlite')
    if not programs:
        programs = ['General']
    
    return campaign_type, ', '.join(programs)

# Process each month
for month in months:
    month_dir = base_dir / month
    
    if not month_dir.exists():
        continue
    
    print(f"\n{'=' * 100}")
    print(f"ANALYZING {month.upper()} 2025")
    print(f"{'=' * 100}")
    
    excel_files = list(month_dir.glob('*.xlsx'))
    excel_files = [f for f in excel_files if not f.name.startswith('~$')]
    
    month_unique_contacts = set()
    month_call_outcomes = {}
    month_interest_levels = {}
    month_delay_reasons = {}
    month_total_calls = 0
    month_reached_calls = 0
    
    print(f"Found {len(excel_files)} campaign files\n")
    
    for file in excel_files:
        try:
            excel_file = pd.ExcelFile(file)
            sheet_names = excel_file.sheet_names
            
            # Look for agent-assigned sheets (skip 'All', 'Cleaned', 'Form Responses', etc.)
            agent_sheets = [s for s in sheet_names if s not in ['All', 'Cleaned', 'Form Responses 1'] 
                          and ('all in one' in s.lower() or 'calls' in s.lower() or len(s) < 15)]
            
            if not agent_sheets:
                continue
            
            campaign_type, program = extract_campaign_info(file.name)
            
            # Read from consolidated sheet first if available
            consolidated_sheet = None
            for sheet in agent_sheets:
                if 'all in one' in sheet.lower() or 'calls' in sheet.lower():
                    consolidated_sheet = sheet
                    break
            
            if consolidated_sheet:
                df = pd.read_excel(file, sheet_name=consolidated_sheet)
            else:
                # Combine all agent sheets
                dfs = []
                for sheet in agent_sheets:
                    try:
                        df_temp = pd.read_excel(file, sheet_name=sheet)
                        dfs.append(df_temp)
                    except:
                        continue
                if not dfs:
                    continue
                df = pd.concat(dfs, ignore_index=True)
            
            # Identify columns (handle variations in column names)
            phone_col = None
            email_col = None
            status_col = None
            interest_col = None
            reason_col = None
            
            for col in df.columns:
                col_lower = str(col).lower().strip()
                if 'phone' in col_lower or 'number' in col_lower:
                    phone_col = col
                if 'email' in col_lower:
                    email_col = col
                if 'call status' in col_lower or 'status' in col_lower:
                    status_col = col
                if 'interested' in col_lower or 'interest' in col_lower:
                    interest_col = col
                if 'reason' in col_lower and 'pause' in col_lower:
                    reason_col = col
            
            if not status_col:
                continue
            
            # Process each contact
            for idx, row in df.iterrows():
                # Track unique contacts
                contact_id = None
                if phone_col and pd.notna(row[phone_col]):
                    contact_id = str(row[phone_col]).strip()
                elif email_col and pd.notna(row[email_col]):
                    contact_id = str(row[email_col]).strip()
                
                if contact_id:
                    month_unique_contacts.add(contact_id)
                    overall_unique_contacts.add(contact_id)
                
                # Count total calls
                month_total_calls += 1
                
                # Categorize call outcome
                call_status = normalize_call_status(row[status_col])
                month_call_outcomes[call_status] = month_call_outcomes.get(call_status, 0) + 1
                overall_call_outcomes[call_status] = overall_call_outcomes.get(call_status, 0) + 1
                
                if call_status == 'Reached':
                    month_reached_calls += 1
                    
                    # Analyze interest level among reached
                    if interest_col:
                        interest = normalize_interest(row[interest_col])
                        month_interest_levels[interest] = month_interest_levels.get(interest, 0) + 1
                        overall_interest_levels[interest] = overall_interest_levels.get(interest, 0) + 1
                    
                    # Analyze delay reasons
                    if reason_col:
                        reason = normalize_delay_reason(row[reason_col])
                        month_delay_reasons[reason] = month_delay_reasons.get(reason, 0) + 1
                        overall_delay_reasons[reason] = overall_delay_reasons.get(reason, 0) + 1
                
                # Store detailed data
                all_contacts_data.append({
                    'Month': month,
                    'Campaign': file.name,
                    'Type': campaign_type,
                    'Program': program,
                    'Contact_ID': contact_id,
                    'Call_Status': call_status,
                    'Interest': normalize_interest(row[interest_col]) if interest_col else 'Unknown',
                    'Delay_Reason': normalize_delay_reason(row[reason_col]) if reason_col else 'Not Specified'
                })
        
        except Exception as e:
            print(f"  ⚠️ Error processing {file.name}: {str(e)}")
            continue
    
    # Store monthly stats
    monthly_detailed_stats[month] = {
        'unique_contacts': len(month_unique_contacts),
        'total_calls': month_total_calls,
        'reached_calls': month_reached_calls,
        'call_outcomes': month_call_outcomes,
        'interest_levels': month_interest_levels,
        'delay_reasons': month_delay_reasons
    }
    
    # Print monthly summary
    if month_total_calls > 0:
        reach_rate = (month_reached_calls / month_total_calls * 100)
        print(f"\n📊 {month} Summary:")
        print(f"  Unique Contacts: {len(month_unique_contacts):,}")
        print(f"  Total Calls Made: {month_total_calls:,}")
        print(f"  Successfully Reached: {month_reached_calls:,} ({reach_rate:.1f}%)")
        
        print(f"\n  Call Outcomes:")
        for outcome, count in sorted(month_call_outcomes.items(), key=lambda x: x[1], reverse=True):
            pct = (count / month_total_calls * 100)
            print(f"    - {outcome}: {count:,} ({pct:.1f}%)")
        
        if month_interest_levels:
            print(f"\n  Interest Levels (Among Reached):")
            for level, count in sorted(month_interest_levels.items(), key=lambda x: x[1], reverse=True):
                pct = (count / month_reached_calls * 100) if month_reached_calls > 0 else 0
                print(f"    - {level}: {count:,} ({pct:.1f}%)")

# Generate comprehensive report
print("\n\n")
print("=" * 100)
print("COMPREHENSIVE SUMMARY REPORT - JANUARY TO AUGUST 2025")
print("=" * 100)

total_calls = sum(stats['total_calls'] for stats in monthly_detailed_stats.values())
total_reached = sum(stats['reached_calls'] for stats in monthly_detailed_stats.values())
total_unique = len(overall_unique_contacts)

print(f"\n📈 OVERALL STATISTICS")
print(f"{'─' * 100}")
print(f"Total Unique Contacts: {total_unique:,}")
print(f"Total Calls Made: {total_calls:,}")
print(f"Total Successfully Reached: {total_reached:,} ({(total_reached/total_calls*100):.1f}%)")
print(f"Average Reach Rate: {(total_reached/total_calls*100):.1f}%")

print(f"\n📞 CALL OUTCOME BREAKDOWN")
print(f"{'─' * 100}")
print(f"{'Outcome':<25} {'Count':>12} {'% of Total':>12}")
print(f"{'─' * 100}")
for outcome, count in sorted(overall_call_outcomes.items(), key=lambda x: x[1], reverse=True):
    pct = (count / total_calls * 100) if total_calls > 0 else 0
    print(f"{outcome:<25} {count:>12,} {pct:>11.1f}%")

print(f"\n💬 INTEREST LEVELS (Among Reached Contacts)")
print(f"{'─' * 100}")
if overall_interest_levels:
    print(f"{'Interest Level':<25} {'Count':>12} {'% of Reached':>15}")
    print(f"{'─' * 100}")
    for level, count in sorted(overall_interest_levels.items(), key=lambda x: x[1], reverse=True):
        pct = (count / total_reached * 100) if total_reached > 0 else 0
        print(f"{level:<25} {count:>12,} {pct:>14.1f}%")

print(f"\n⏸️  REASONS FOR DELAYED ENROLLMENT")
print(f"{'─' * 100}")
if overall_delay_reasons:
    reasons_with_data = {k: v for k, v in overall_delay_reasons.items() if k != 'Not Specified'}
    total_with_reasons = sum(reasons_with_data.values())
    
    if total_with_reasons > 0:
        print(f"{'Reason':<30} {'Count':>12} {'% of Those with Delays':>25}")
        print(f"{'─' * 100}")
        for reason, count in sorted(reasons_with_data.items(), key=lambda x: x[1], reverse=True):
            pct = (count / total_with_reasons * 100)
            print(f"{reason:<30} {count:>12,} {pct:>24.1f}%")

# Generate insights
print(f"\n💡 KEY INSIGHTS")
print(f"{'─' * 100}")

if total_calls > 0:
    reach_pct = (total_reached / total_calls * 100)
    print(f"\n1. REACH EFFECTIVENESS:")
    print(f"   Out of a total of {total_calls:,} calls, {reach_pct:.1f}% were successfully reached and spoken to.")
    
    if 'Interested' in overall_interest_levels and 'Considering' in overall_interest_levels:
        interested_pct = (overall_interest_levels.get('Interested', 0) / total_reached * 100)
        considering_pct = (overall_interest_levels.get('Considering', 0) / total_reached * 100)
        print(f"\n2. ENGAGEMENT QUALITY:")
        print(f"   Among those reached, {interested_pct:.1f}% expressed continued interest,")
        print(f"   while {considering_pct:.1f}% said they were still considering their options.")
    
    if reasons_with_data and len(reasons_with_data) >= 3:
        sorted_reasons = sorted(reasons_with_data.items(), key=lambda x: x[1], reverse=True)
        reason1, count1 = sorted_reasons[0]
        reason2, count2 = sorted_reasons[1]
        reason3, count3 = sorted_reasons[2]
        
        pct1 = (count1 / total_with_reasons * 100)
        pct2 = (count2 / total_with_reasons * 100)
        pct3 = (count3 / total_with_reasons * 100)
        
        print(f"\n3. ENROLLMENT DELAYS:")
        print(f"   The leading reason for delaying enrollment was {reason1.lower()} ({pct1:.1f}%).")
        print(f"   The second most common reason was {reason2.lower()} ({pct2:.1f}%),")
        print(f"   followed by {reason3.lower()} ({pct3:.1f}%).")

# Monthly breakdown
print(f"\n📅 MONTHLY DETAILED BREAKDOWN")
print(f"{'─' * 100}")
print(f"{'Month':<12} {'Unique':>10} {'Total Calls':>12} {'Reached':>10} {'Reach %':>10} {'Interested %':>15}")
print(f"{'─' * 100}")
for month in months:
    if month in monthly_detailed_stats:
        stats = monthly_detailed_stats[month]
        reach_pct = (stats['reached_calls'] / stats['total_calls'] * 100) if stats['total_calls'] > 0 else 0
        interested = stats['interest_levels'].get('Interested', 0)
        interested_pct = (interested / stats['reached_calls'] * 100) if stats['reached_calls'] > 0 else 0
        
        print(f"{month:<12} {stats['unique_contacts']:>10,} {stats['total_calls']:>12,} {stats['reached_calls']:>10,} {reach_pct:>9.1f}% {interested_pct:>14.1f}%")

# Save detailed data
df_all = pd.DataFrame(all_contacts_data)
output_file = base_dir / 'detailed_call_outcomes_analysis.csv'
df_all.to_csv(output_file, index=False)

print(f"\n✅ Detailed call outcomes data saved to: {output_file}")
print(f"\n{'=' * 100}")
print(f"Report generated on: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
print(f"{'=' * 100}")
