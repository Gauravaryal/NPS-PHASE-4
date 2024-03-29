from ReconcileAbstract import Reconcile
from qrlib.QRUtils import display
from Constants import Banks_FILEPATH
from Constants import SOA_FILEPATH
from xlsxwriter import Workbook
from utils.Utils import run_phase
import pandas as pd
import re

class SaptaKoshiBank(Reconcile):
    def __init__(self, bank_name, excelwriter):
        """
        Parameters:
            bank_name (str): Name of the bank.
            excelwriter: Excel writer object for generating reports.
        """
        super().__init__(bank_name, excelwriter)
        self.bank_statement_df = pd.read_excel(Banks_FILEPATH.SAPTA_KOSHI_BANK_FILE, skiprows= 1)
        self.soa_statement_df = pd.read_csv(SOA_FILEPATH.SAPTA_KOSHI_SOA, encoding='latin-1')
        try:
            display(self.soa_statement_df.loc[0, 'Transaction Id'])
        except:
            self.soa_statement_df = pd.read_csv(SOA_FILEPATH.SAPTA_KOSHI_SOA, encoding='latin-1', skiprows=1)
        

    def main(self):
        '''
         This method orchestrates the entire reconciliation process by calling various steps.
        '''
        display(f'{self.bank_statement_df.head()}')
        display(f'{self.soa_statement_df.head()}')
        self.preprocessing_bank_stmt()
        self.updated_standardize_bank_stmt()
        self.extracting_tid_from_bank_stmt()
        self.preprocessing_soa_stmt()
        self.matching_bank_stmt_with_soa_report()
        self.matching_soa_report_with_bank_stmt()
        self.total_debit_credit_amount_of_soa()
        self.total_debit_credit_amount_of_bank()
        self.debit_credit_amount_matches_tid_of_bank_with_soa()
        self.debit_credit_amount_of_soa_matched_with_bank()
        self.extract_number_of_tid_CR_and_DR_from_soa()
        self.extract_number_of_tid_CR_and_DR_from_bank_stmt()
        self.write_soa_data()
        self.write_bank_data()
        self.generate_report()

    @run_phase(phase_number=1)
    def preprocessing_bank_stmt(self):
        '''
         This method cleans and formats the bank statement DataFrame and also format the Date to match it with the Soa report
        '''
        self.bank_statement_df.columns = self.bank_statement_df.columns.str.strip()
        self.bank_statement_df['Date'] = pd.to_datetime(self.bank_statement_df['TranDate'], format="%d/%m/%Y").dt.date
        display(f'preprocessed_bank_stmt--> {self.bank_statement_df.head()}')
        display(f'columns-->{self.bank_statement_df.columns}')  
    
    @run_phase(phase_number=1)
    def updated_standardize_bank_stmt(self):
        '''
        applies formatting to the column by calling the Standard_format function.
        '''
        self.bank_statement_df[['Mode', 'Amount']] = self.bank_statement_df.apply(self.standard_format, axis=1)    

    @run_phase(phase_number=1)
    def extracting_tid_from_bank_stmt(self):
        """
        Extraccting the transaction id from the column 'Desc2' using regular expression 
        """
        for index, row in self.bank_statement_df.iterrows():
            if "100000" in str(row["Desc2"]):
                id_matches = re.findall(r'10*[0-9]{6}', str(row["Desc2"]))
                if id_matches:
                    id = id_matches[0].replace("100000", "")
                    self.bank_statement_df.loc[index, 'Transaction Id'] = id

            elif "10000" in str(row["Desc2"]):
                id_matches = re.findall(r'10*[0-9]{7}', str(row["Desc2"]))
                if id_matches:
                    id = id_matches[0].replace("10000", "")
                    self.bank_statement_df.loc[index, 'Transaction Id'] = id
        display(f'total_tid {self.bank_statement_df["Transaction Id"].count()}')

    @run_phase(phase_number=1)
    def standard_format(self,row):
        """
        Determine transaction mode and amount from bank statement data.

        Parameters:
            row (Series): A row of bank statement data.

        Returns:
            Series: Mode (CR/DR) and amount.
        """
        if row['Amount'] < 0:
            mode = 'DR'
            amount = abs(row['Amount'])
        elif row['Amount'] > 0:
            mode = 'CR'
            amount = abs(row['Amount'])
        else:
            mode = None
            amount = None
        return pd.Series([mode, amount], index=['Mode', 'Amount'])


    

