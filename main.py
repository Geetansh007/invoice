"""
Excel Record Processor

This module provides functionality to process Excel records, combining records with
the same email address and dates from the same month/year.
"""
import pandas as pd
from datetime import datetime
from collections import defaultdict
import num2words
from invoice_generato import InvoiceDocGeneratorXML
from mail_invoice import send_invoices_for_records
import getpass


class RecordProcessor:
    """
    A class to process and combine Excel records based on specific criteria.
    
    This class provides methods to read Excel files, process the data,
    combine records that share the same email address and have dates in the 
    same month and year, while preserving all other fields.
    """
    
    def __init__(self):
        """Initialize the RecordProcessor with empty data structures."""
        self.original_records = [] 
        self.processed_records = []  
        self.invalid_records = []   
    
    def read_excel_file(self, file_path):
        """
        Read data from an Excel file.
        
        Args:
            file_path (str): Path to the Excel file
            
        Returns:
            list: List of dictionary records from the Excel file
        """
        df = pd.read_excel(file_path)
        all_records = df.to_dict(orient='records')
        self.original_records = []
        self.invalid_records = []
        
        for record in all_records:
            if self._validate_record(record):
                date = record.get('Date of Entry')
                if isinstance(date, str):
                    date = datetime.strptime(date, '%Y-%m-%d').date()
                record['Date of Entry'] = date.date()
                self.original_records.append(record)
            else:
                self.invalid_records.append(record)
                
    
    def _validate_record(self, record):
        """
        Validate a record by checking for null or N/A values in required fields.
        Address field is allowed to be empty.
        
        Args:
            record (dict): The record to validate
            
        Returns:
            bool: True if the record is valid, False otherwise
        """
        required_fields = [
            'Date of Entry', 'Amount', 'Description', 'Beneficiary Name', 
            'Bank Name', 'Bank Account No', 'IFSC/SWIFT Code', 'email'
        ]
        for field in required_fields:
            if field not in record:
                continue
                
            value = record.get(field)
            
           
            if pd.isna(value): 
                return False
            if isinstance(value, str):
                value_str = value.strip().upper()
                if value_str == '' or value_str == 'N/A' or value_str == 'NA' or value_str == 'NULL':
                    return False
        return True
    
    def process_records(self):
        """
        Process the records to combine those with the same email and month/year.
        Only Amount and Description fields are combined into lists.
        All other fields are preserved from the first record in each group.
        
        Returns:
            list: Processed records with appropriate records combined
        """
        self.processed_records = self._combine_records(self.original_records)
        self.processed_records = self._add_invoice_details(self.processed_records)
        return self.processed_records
    
    def _combine_records(self, data_list):
        """
        Combine records with the same email and date (month/year).
        This is a private helper method used by process_records.
        
        Args:
            data_list (list): List of record dictionaries to process
            
        Returns:
            list: List of records with appropriate records combined
        """
      
        groups = self._group_records_by_email_and_date(data_list)
        result_records, combined_indices = self._process_groups(groups)
        
        for i, record in enumerate(data_list):
            if i not in combined_indices:
                result_records.append(record)
        
        return result_records
    
    def _group_records_by_email_and_date(self, data_list):
        """
        Group records by email and month/year of the date.
        
        Args:
            data_list (list): List of record dictionaries
            
        Returns:
            defaultdict: Dictionary with (email, month_year) keys and lists of record tuples as values
        """
        groups = defaultdict(list)
        
        for i, record in enumerate(data_list):
            if not record.get('email') or pd.isna(record.get('email')) or str(record.get('email')).strip() == '':
                continue
                
            date = record.get('Date of Entry')
            if isinstance(date, str):
                date = datetime.strptime(date, '%Y-%m-%d')
            month_year = None
            if hasattr(date, 'month') and hasattr(date, 'year'):
                month_year = (date.month, date.year)
            
            if not month_year:
                continue
            email = str(record['email']).strip()
            key = (email, month_year)
            
            groups[key].append((i, record))
        
        return groups
    
    def _process_groups(self, groups):
        """
        Process each group to combine records if needed.
        
        Args:
            groups (defaultdict): Dictionary with (email, month_year) keys and lists of record tuples
            
        Returns:
            tuple: (result_records, combined_indices) where:
                - result_records is a list of processed records
                - combined_indices is a set of indices of records that were combined
        """
        result_records = []
        combined_indices = set()
        for (email, month_year), group in groups.items():
            if len(group) > 1:
                base_record = group[0][1].copy()
                amounts = []
                descriptions = []
                for idx, record in group:
                    amounts.append(record['Amount'])
                    descriptions.append(record['Description'])
                    combined_indices.add(idx)
                base_record['Amount'] = amounts
                base_record['Description'] = descriptions
                result_records.append(base_record)
        
        return result_records, combined_indices
    
    "ADDING INVOICE DETAILS"
    
    def _add_invoice_details(self, records):
        month_year_counts = defaultdict(int)
        
        for record in records:
            date = record.get('Date of Entry')
            if isinstance(date, str):
                date = datetime.strptime(date, '%Y-%m-%d')
            
            month_abbr = date.strftime('%b').lower()
            year = date.strftime('%Y')
            month_year = f"{month_abbr}{year}"
            
            month_year_counts[month_year] += 1
            count = month_year_counts[month_year]
            
            invoice_number = f"{month_year}{count:03d}"
            record['Invoice Number'] = invoice_number
            
            if isinstance(record['Amount'], list):
                total_amount = sum(record['Amount'])
            else:
                total_amount = record['Amount']
            
            record['Total Amount'] = total_amount
            
            try:
                dollars = int(total_amount)
                cents = int(round((total_amount - dollars) * 100))
                if dollars > 0:
                    dollars_in_words = num2words.num2words(dollars)
                    dollars_text = f"{dollars_in_words} dollar"
                    if dollars != 1:
                        dollars_text += "s"
                else:
                    dollars_text = "Zero dollars"
                if cents > 0:
                    cents_in_words = num2words.num2words(cents)
                    cents_text = f"{cents_in_words} cent"
                    if cents != 1:
                        cents_text += "s"
                    amount_in_words = f"{dollars_text} and {cents_text}"
                else:
                    amount_in_words = dollars_text
                
                record['Total Amount in Words'] = amount_in_words.capitalize()
            except:
                record['Total Amount in Words'] = f"{total_amount} dollars"
        
        return records
    
    def display_results(self):
        """
        Display the results showing original and processed records.
        """
        print("\nORIGINAL RECORDS:")
        for i, record in enumerate(self.original_records):
            print(f"{i+1}. Amount: {record.get('Amount')}, "
                  f"Description: {record.get('Description')}, "
                  f"Email: {record.get('email')}, "
                  f"Date: {record.get('Date of Entry')}")
        
        if self.invalid_records:
            print(f"\nINVALID RECORDS (FILTERED OUT): {len(self.invalid_records)}")
            for i, record in enumerate(self.invalid_records):
                print(f"{i+1}. Invalid record: "
                      f"Email: {record.get('email')}, "
                      f"Date: {record.get('Date of Entry')}, "
                      f"Amount: {record.get('Amount')}, "
                      f"Description: {record.get('Description')}, "
                      f"Bank Name: {record.get('Bank Name')}, "
                      f"IFSC/SWIFT Code: {record.get('IFSC/SWIFT Code')}")
        
        print("\nFINAL PROCESSED RECORDS (WITH ALL FIELDS):")
        for i, record in enumerate(self.processed_records):
            print(f"\n{'-' * 60}")
            print(f"RECORD {i+1}:")
            print(f"Invoice Number: {record.get('Invoice Number')}")
            print(f"Date of Entry: {record.get('Date of Entry')}")
            print(f"Email: {record.get('email')}")
            
            if isinstance(record.get('Amount'), list):
                print(f"Combined Amounts: {record.get('Amount')}")
                print(f"Combined Descriptions: {record.get('Description')}")
            else:
                print(f"Amount: {record.get('Amount')}")
                print(f"Description: {record.get('Description')}")
                
            print(f"Total Amount: {record.get('Total Amount')}")
            print(f"Total Amount in Words: {record.get('Total Amount in Words')}")
            print(f"Beneficiary Name: {record.get('Beneficiary Name')}")
            print(f"Bank Name: {record.get('Bank Name')}")
            print(f"Bank Account No: {record.get('Bank Account No')}")
            print(f"IFSC/SWIFT Code: {record.get('IFSC/SWIFT Code')}")
            print(f"IBAN No: {record.get('IBAN No')}")
            print(f"Address: {record.get('Address')}")
            print(f"{'-' * 60}")


def main():
    """
    Main function to demonstrate the RecordProcessor.
    """
    try:
        processor = RecordProcessor()
        print("\n" + "="*50)
        print("Processing main Excel file...")
        processor.read_excel_file("Dummy details_styled.xlsx")
        processed_records = processor.process_records()
        processor.display_results()
        print("\n" + "="*50)
        print("Generating invoices...")
        invoice_generator = InvoiceDocGeneratorXML(
            template_path="Invoice format.docx",
            output_dir="generated_invoices",
            data_list=processed_records
        )
        invoice_generator.generate_documents()
        # Email functionality
        sender_email = "geetanshjoshi007@gmail.com"
        sender_password = getpass.getpass("Enter sender email password: ")
        send_invoices_for_records(
            processed_records,
            "generated_invoices",
            sender_email,
            sender_password
        )
    except FileNotFoundError:
        print("\nMain data file not found.")


if __name__ == "__main__":
    main()

