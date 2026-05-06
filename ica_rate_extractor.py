import win32com.client
import os
import tempfile
import re
import pandas as pd
from pdfminer.high_level import extract_text
from datetime import datetime
from openpyxl import load_workbook


def connect_to_outlook():
    """Connect to Outlook and return the namespace."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return namespace


def get_emails_with_ica_pdfs(namespace, start_date):
    """Find emails in Inbox from @StoneCastle.com with PDF attachments containing 'ICA' in filename."""
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
    
    matching_emails = []
    
    # Filter emails from @StoneCastle.com using Restrict for faster search
    filter_str = "@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%@stonecastle.com'"
    filtered_items = inbox.Items.Restrict(filter_str)
    
    for item in filtered_items:
        try:
            # Check if email was received on or after start_date
            received_time = item.ReceivedTime
            email_date = datetime(received_time.year, received_time.month, received_time.day)
            
            if email_date < start_date:
                continue
            
            # Check for PDF attachments with "ICA" in the name
            if item.Attachments.Count > 0:
                for attachment in item.Attachments:
                    if attachment.FileName.lower().endswith('.pdf') and 'ica' in attachment.FileName.lower():
                        matching_emails.append({
                            'email': item,
                            'attachment': attachment,
                            'received_date': email_date
                        })
        except Exception as e:
            continue
    
    return matching_emails


def extract_pdf_attachment(attachment, temp_dir):
    """Save PDF attachment to temp directory and return path."""
    pdf_path = os.path.join(temp_dir, attachment.FileName)
    attachment.SaveAsFile(pdf_path)
    return pdf_path


def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file using pdfminer."""
    return extract_text(pdf_path)


def parse_ica_rates(text, received_date):
    """
    Parse ICA rate data from extracted PDF text.
    Returns a list of dictionaries with Date, Institution, APY, APR.
    Returns empty list if required headers are not found.
    """
    # Check if this PDF contains the required headers
    has_institution = bool(re.search(r'\bInstitution\b', text, re.IGNORECASE))
    has_apy = bool(re.search(r'\bAPY\b', text, re.IGNORECASE))
    has_apr = bool(re.search(r'\bAPR\b', text, re.IGNORECASE))
    
    if not (has_institution and has_apy and has_apr):
        return []
    
    results = []
    
    # Split text into lines for parsing
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    # Skip header words
    skip_words = ['institution', 'apy', 'apr', 'daily', 'rates', 'ica']
    
    # Collect institutions (text lines) and rates (lines with percentages)
    institutions = []
    rate_pairs = []
    
    for line in lines:
        # Skip header line or footer content
        if any(word in line.lower() for word in ['further information', 'please visit', 'customer service', 
                                                   'investment advice', 'solicitation', 'current yield',
                                                   'guarantee', 'https://', 'close of the day']):
            break
        
        # Check if line contains rate values (percentages)
        rate_match = re.findall(r'(\d+\.\d+)%?', line)
        
        if len(rate_match) >= 2:
            # This line has APY and APR values
            # Check if there's also an institution name on this line
            inst_match = re.match(r'^([A-Za-z][A-Za-z\s\.\,\&\-]+?)\s+\d+\.\d+', line)
            if inst_match:
                inst_name = inst_match.group(1).strip()
                # Skip if it's a header word
                if inst_name.lower() not in skip_words:
                    institutions.append(inst_name)
            
            # Store the rate pair
            rate_pairs.append((float(rate_match[0]), float(rate_match[1])))
        
        elif re.match(r'^[A-Za-z]', line) and not any(skip in line.lower() for skip in skip_words):
            # This is an institution name line (starts with letter, no rates)
            # Clean up the institution name
            inst_name = line.strip()
            if inst_name and len(inst_name) > 2:
                institutions.append(inst_name)
    
    # Match institutions with rate pairs
    if len(institutions) == len(rate_pairs):
        for i, inst in enumerate(institutions):
            apy_val, apr_val = rate_pairs[i]
            results.append({
                'Date': received_date.strftime('%Y-%m-%d'),
                'Institution': inst,
                'APY': apy_val,
                'APR': apr_val
            })
    
    return results


def load_existing_data(filepath):
    """Load existing data from Excel file if it exists."""
    if os.path.exists(filepath):
        try:
            df = pd.read_excel(filepath)
            # Ensure Date column is string format for comparison
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')
            return df
        except Exception as e:
            print(f"Warning: Could not read existing file: {e}")
            return pd.DataFrame(columns=['Date', 'Institution', 'APY', 'APR'])
    return pd.DataFrame(columns=['Date', 'Institution', 'APY', 'APR'])


def save_data(df, filepath):
    """Save DataFrame to Excel with formatting."""
    # Ensure directory exists
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ICA Rates', index=False)
        
        # Apply number format to APY and APR columns
        ws = writer.sheets['ICA Rates']
        for col_name in ['APY', 'APR']:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name) + 1
                for row in range(2, len(df) + 2):
                    ws.cell(row=row, column=col_idx).number_format = '0.00'


def get_start_date():
    """Prompt user for start date."""
    while True:
        date_input = input("\nEnter the starting date (YYYY-MM-DD) or press Enter for all dates: ").strip()
        
        if not date_input:
            return datetime(1900, 1, 1)  # Return very old date to get all emails
        
        try:
            return datetime.strptime(date_input, '%Y-%m-%d')
        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD format.")


def main():
    """Main function to extract ICA rates from Outlook emails."""
    print("=" * 60)
    print("ICA Rate Extractor")
    print("=" * 60)
    
    # Get start date from user
    start_date = get_start_date()
    print(f"\nSearching for emails from {start_date.strftime('%Y-%m-%d')} onwards...")
    
    # Output file path
    output_filepath = r"C:\Users\rfarahmand\Box\Management Company\FinOps\Treasury\Investments\StoneCastle\ICA Historical Rates.xlsx"
    
    # Load existing data
    existing_df = load_existing_data(output_filepath)
    existing_dates = set(existing_df['Date'].unique()) if not existing_df.empty else set()
    print(f"Existing records: {len(existing_df)} rows, {len(existing_dates)} unique dates")
    
    # Connect to Outlook
    print("\nConnecting to Outlook...")
    namespace = connect_to_outlook()
    
    # Find matching emails
    print("Searching for emails with ICA PDF attachments...")
    matching_emails = get_emails_with_ica_pdfs(namespace, start_date)
    print(f"Found {len(matching_emails)} matching email(s) with ICA PDFs")
    
    if not matching_emails:
        print("\nNo matching emails found.")
        return
    
    # Extract data from PDFs
    all_new_data = []
    
    with tempfile.TemporaryDirectory() as temp_dir:
        for item in matching_emails:
            email = item['email']
            attachment = item['attachment']
            received_date = item['received_date']
            date_str = received_date.strftime('%Y-%m-%d')
            
            # Skip if we already have data for this date
            if date_str in existing_dates:
                print(f"  Skipping {attachment.FileName} - data for {date_str} already exists")
                continue
            
            print(f"\nProcessing: {attachment.FileName} (received {date_str})")
            
            try:
                # Extract PDF
                pdf_path = extract_pdf_attachment(attachment, temp_dir)
                
                # Extract text
                text = extract_text_from_pdf(pdf_path)
                
                # Parse rates
                rates = parse_ica_rates(text, received_date)
                
                if rates:
                    print(f"  Extracted {len(rates)} institution(s)")
                    all_new_data.extend(rates)
                else:
                    print(f"  No valid ICA rate data found (missing required headers)")
                    
            except Exception as e:
                print(f"  Error processing: {e}")
    
    # Combine with existing data
    if all_new_data:
        new_df = pd.DataFrame(all_new_data)
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        
        # Sort by Date descending, then Institution
        combined_df['Date'] = pd.to_datetime(combined_df['Date'])
        combined_df = combined_df.sort_values(['Date', 'Institution'], ascending=[False, True])
        combined_df['Date'] = combined_df['Date'].dt.strftime('%Y-%m-%d')
        
        # Save to Excel
        save_data(combined_df, output_filepath)
        
        print("\n" + "=" * 60)
        print(f"SUMMARY")
        print("=" * 60)
        print(f"New records added: {len(all_new_data)}")
        print(f"Total records: {len(combined_df)}")
        print(f"\nResults saved to: {output_filepath}")
    else:
        print("\n" + "=" * 60)
        print("No new data to add.")
        print("=" * 60)


if __name__ == "__main__":
    main()
