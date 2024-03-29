from RPA.Email.ImapSmtp import ImapSmtp
from robot.libraries.BuiltIn import BuiltIn

# from qrlib.QRComponent import QRComponent
from qrlib.QREnv import QREnv
from qrlib.QRUtils import get_secret
from Constants import *
from utils.Utils import get_report_file_path
from qrlib.QRUtils import display
import datetime

# class Email:
#     def __init__(self):
#         pass
        
def send_email():
    current_directory = os.getcwd()

    parent_directory = os.path.abspath(os.path.join(current_directory, os.pardir))
    today = str(datetime.datetime.date(datetime.datetime.now()))
    fll_path = os.path.join(parent_directory, today)
    
    today_date = str(datetime.datetime.today() - timedelta(days=1)).split()[0]
    yesterday_date = today_date.replace("-", "_")
    
    email = ImapSmtp()
    display('inside email')
    print('inside email')
    email.authorize_smtp(account="reconciliation@nepalpayment.com", password="x3Ww68b.xDv;42n#", smtp_server="smtp-mail.outlook.com", smtp_port=587)
    email.send_message(
        sender="reconciliation@nepalpayment.com",
        recipients="support@nepalpayment.com",
        # cc="reconciliation@nepalpayment.com",
        subject=f"Report of {yesterday_date}",
        body=f"""        
    This report as of {yesterday_date} is generated by bot.

    Do not reply.
    """,
    attachments=[f"{fll_path}\\Reconciliation_Matched_report_{yesterday_date}.xlsx", f"{fll_path}\\Reconciliation_Unmatched_report_{yesterday_date}.xlsx"]
        )

# class Email(QRComponent):
#     def __init__(self):
#         # send mail values
#         super().__init__()
#         self.sending_file_list = []
#         self.subject = ''
#         self.body = ''
#         self.recipients = []
        
#         # smtp connection
#         self.account = ''
#         self.password = ''
#         self.server = ''
#         self.port = ''

#         self.__vault_data = ''

#         self.email_recievers=[]
        
#         self.__recievers_data=''

#     def _get_vault(self):
#         logger = self.logger
#         self.__vault_data = get_secret('smtp')
#         self.__recievers_data = get_secret('recievers')
    
#     def _set_smtp_creds(self):
#         self.account = self.__vault_data['account']
#         self.server = self.__vault_data['server']
#         self.port = self.__vault_data['port']
#         try:
#             self.password = self.__vault_data['password']
#         except:
#             self.password = None

#     def _get_reciepents(self):
#         self.recipients = self.__recievers_data['email'].split(",")

    
#     def _authmail_and_send(self):
#         """Call when send mail only"""
#         logger = self.logger
#         mail = ImapSmtp(smtp_server=self.server, smtp_port=self.port)
#         mail.authorize(account=self.account, password=self.password)
#         mail.send_message(
#             sender=self.account,
#             recipients=self.recipients,
#             subject=self.subject,
#             body=self.body,
#             attachments=self.sending_file_list,
#             html=True
#         )
#         logger.info("Mail Sent Successfully.")
        

#     def initiate_connection(self):
#         self._get_vault()
#         self._set_smtp_creds()
#         self._get_reciepents()


#     def send_summary_mail(self):
#         try:
#             self.initiate_connection()
#             actual_revised_excel_path = get_report_file_path()
#             self.sending_file_list.append(actual_revised_excel_path)
#             self.subject = f"NEWS PORTAL Report for QUICK SAMACHAR Details."
#             self.body = f'''
#                 Dear All,<br><br>\
#                     <p>Please find the NEWS PORTAL data: </p>          
                
#                     <p>This mail is auto generated by the system. Please do not reply.</p><br><br>
#                     <strong>Thank you!!!</strong>
#             '''
#             self._authmail_and_send()
#         except Exception as e:
#             display(f'unable to send mail because of {e}')
