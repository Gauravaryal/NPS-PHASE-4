from ReconcileAbstract import Reconcile
from qrlib.QRUtils import display
from Constants import Banks_FILEPATH
from Constants import SOA_FILEPATH
from xlsxwriter import Workbook
from utils.Utils import run_phase
import warnings
warnings.filterwarnings("ignore")
warnings.resetwarnings()
import pandas as pd
import re

class BestFinance(Reconcile):
    def __init__(self, bank_name, excelwriter):
        """
        Parameters:
            bank_name (str): Name of the bank.
            excelwriter: Excel writer object for generating reports.
        """
        super().__init__(bank_name, excelwriter)
        self.bank_statement_df = pd.read_excel(Banks_FILEPATH.BESTFINANCE_BANK_FILE,skiprows=1)
        self.soa_statement_df = pd.read_csv(SOA_FILEPATH.BESTFINANCE_SOA_FILE, encoding='latin-1')
        try:
            display(self.soa_statement_df.loc[0, 'Transaction Id'])
        except:
            self.soa_statement_df = pd.read_csv(SOA_FILEPATH.BESTFINANCE_SOA_FILE, encoding='latin-1', skiprows=1)       

    def main(self):
        '''
         This method orchestrates the entire reconciliation process by calling various steps.
        '''
        # self.preprocessing_bank_stmt()
        # self.preprocessing_bank_stmt_phase4()
        # self.updated_standardize_bank_stmt()
        self.updated_standardize_bank_stmt_phase_4()
        # self.extracting_tid_from_bank_stmt()
        self.extracting_tid_from_bank_stmt_phase_4()
        display(f'asdf {self.bank_statement_df.columns}')
        
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


    @run_phase(phase_number=4)
    def preprocessing_bank_stmt_phase_4(self):
        '''
         This method cleans and formats the bank statement DataFrame and also format the Date to match it with the Soa report if needed
        '''
        self.bank_statement_df['Date'] = pd.to_datetime(self.bank_statement_df['Date'], format="%Y-%m-%dT%H:%M:%S").dt.date


    @run_phase(phase_number=4)     
    def updated_standardize_bank_stmt_phase_4(self):
        '''
        applies formatting to columns.
        '''
        self.bank_statement_df[['Mode', 'Amount']] = self.bank_statement_df.apply(self.standard_format_phase_4, axis=1)    


    @run_phase(phase_number=4)
    def extracting_tid_from_bank_stmt_phase_4(self):
        """
        Extract transaction IDs from bank statement data.
        This method extracts transaction IDs from "Desc3" columns.
        """
        display('aaye ta bhitra')
        for index, row in self.bank_statement_df.iterrows():
            for roww in (str(row['Desc3']), str(row['Desc2']), str(row['Desc1'])):
                if "NPS-IF-" in roww:
                    ind = roww.index('NPS-IF-')
                    id = roww[ind+7:ind+14]
                    self.bank_statement_df.loc[index, 'Transaction Id'] = id[0]
                    
                elif "FTMS-" in roww:
                    ind = roww.index('FTMS-')
                    id = roww[ind+5:ind+11]
                    self.bank_statement_df.loc[index, 'Transaction Id'] = id

                elif "10000" in roww:
                    ind = roww.index('10000')
                    id = roww[ind+5:ind+12]
                    self.bank_statement_df.loc[index, 'Transaction Id'] = id


    @run_phase(phase_number=4)
    def standard_format_phase_4(self,row):
        """
        Determine transaction mode and amount from bank statement data.
        Parameters:
            row (Series): A row of bank statement data.
        Returns:
            Series: Mode (CR/DR) and amount.
        """
        if row['Txn Type'] == 'DR' :
            mode = 'DR'
            amount = abs(row['Amount'])
        elif row['Txn Type'] == 'CR':
            mode = 'CR'
            amount = abs(row['Amount'])
        else:
            mode = None
            amount = None
        return pd.Series([mode, amount], index=['Mode', 'Amount'])
