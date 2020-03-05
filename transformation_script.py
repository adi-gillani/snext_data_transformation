import pandas as pd
from datetime import datetime, timedelta
import yagmail
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('mode.chained_assignment', None)

#listing down correct column names for different data files
customer_data_column_names = ['store_id', 'customer_manager_code', 'customer_id', 'company_name', 'customer_name',
                              'phone_number']
invoice_data_column_names = ['store_id', 'store_name', 'segment', 'store_id1', 'store_name1', 'customer_id',
                             'customer_name', 'cu_attention1', 'cu_attention2', 'cu_attention3', 'invoice_timestamp',
                             'cu_invoice1', 'invoice_id', 'invoice_id_identifier', 'settlement_cd1', 'payment_method',
                             'total']
final_sheet_column_names = ['Store ID', 'Manager Code', 'Unique ID', 'Company', 'Contact Person', 'Contact Number', 'Secondary Division', 
                            'Transaction Timestamp', 'Payment Method', 'Total', 'Invoice ID', 'Primary Division', 'Medium', 'Feedback Group']

#importing customer data and standardizing column names for B2B FSD
fsd_customer_data = pd.read_excel("customer_list_b2b_fsd.xlsx")
fsd_customer_data.columns = customer_data_column_names

#importing invoices data and standardizing columns names for B2B FSD - Horeca and O&I
invoice_data_horeca_fsd = pd.read_excel("NPS Horeca FSD.xlsx", skiprows=2)
invoice_data_horeca_fsd.columns = invoice_data_column_names 

invoice_data_oi_fsd = pd.read_excel("NPS O&I FSD.xlsx", skiprows=2)
invoice_data_oi_fsd.columns = invoice_data_column_names

#binding both O&I and Horeca data into one sheet
invoice_data_fsd = invoice_data_horeca_fsd.append(invoice_data_oi_fsd)

#merging invoice data and customer data 
transformed_sheet_fsd = pd.merge(left=fsd_customer_data, right=invoice_data_fsd, left_on=['store_id', 'customer_id'], right_on=['store_id', 'customer_id'])

#adding unique invoice id
transformed_sheet_fsd['unique_invoice_id'] = transformed_sheet_fsd['invoice_id_identifier'].astype(str) + transformed_sheet_fsd['invoice_id'].astype(str)

#removing unwanted columns
final_sheet_fsd = transformed_sheet_fsd.iloc[:,[0,1,2,3,4,5,7,14,19,20,21]]

#adding column for primary division
final_sheet_fsd.loc[final_sheet_fsd['store_id'] == 10, 'primary_division'] = "Thokar Niaz Baig"
final_sheet_fsd.loc[final_sheet_fsd['store_id'] == 11, 'primary_division'] = "Capital"
final_sheet_fsd.loc[final_sheet_fsd['store_id'] == 12, 'primary_division'] = "Safari"
final_sheet_fsd.loc[final_sheet_fsd['store_id'] == 15, 'primary_division'] = "Faisalabad"
final_sheet_fsd.loc[final_sheet_fsd['store_id'] == 16, 'primary_division'] = "DHA Store"
final_sheet_fsd.loc[final_sheet_fsd['store_id'] == 22, 'primary_division'] = "Manghopir"
final_sheet_fsd.loc[final_sheet_fsd['store_id'] == 23, 'primary_division'] = "Star Gate"
final_sheet_fsd.loc[final_sheet_fsd['store_id'] == 25, 'primary_division'] = "Ravi Road"
final_sheet_fsd.loc[final_sheet_fsd['store_id'] == 26, 'primary_division'] = "Model Town"

#adding column for medium
final_sheet_fsd['medium'] = 'SMS'

#adding column for feedback group type
final_sheet_fsd['type'] = 'FSD'

#converting segment from O&I to Office and Industry
final_sheet_fsd['segment'].loc[(final_sheet_fsd['segment'] == 'O&I')] = "Office and Industry"

#standardizing column names for final FSD sheet
final_sheet_fsd.columns = final_sheet_column_names

#exporting CSV
final_sheet_fsd.to_csv("B2B FSD.csv", index=False)

###########################################################
###########DOING THE SAME SHIT FOR B2B IN-STORE############
###########################################################

#importing customer data for B2B In-Store and applying standarized column names
customer_data_b2b_in_store = pd.read_csv('Book1.csv', engine='python')
customer_data_b2b_in_store.columns = customer_data_column_names

#importing invoice data and standardizing column names for B2B In-Store
b2b_in_store_horeca = pd.read_excel('NPS Horeca B2B.xlsx', skiprows=2)
b2b_in_store_horeca.columns = invoice_data_column_names

b2b_in_store_oi = pd.read_excel('NPS O&I B2B.xlsx', skiprows=2)
b2b_in_store_oi.columns = invoice_data_column_names

#binding O&I and Horeca into one sheet
invoice_data_b2b_in_store = b2b_in_store_horeca.append(b2b_in_store_oi)

#merging customer and invoice data for B2B In-Store
transformed_sheet_b2b_in_store = pd.merge(left=customer_data_b2b_in_store, right=invoice_data_b2b_in_store, left_on=['store_id', 'customer_id'], right_on=['store_id', 'customer_id'])

#adding unique invoice ID
transformed_sheet_b2b_in_store['unique_invoice_id'] = transformed_sheet_b2b_in_store['invoice_id_identifier'].astype(str) + transformed_sheet_b2b_in_store['invoice_id'].astype(str)

#removing unwanted columns
final_sheet_b2b_in_store = transformed_sheet_b2b_in_store.iloc[:,[0,1,2,3,4,5,7,14,19,20,21]]

#adding column for primary division
final_sheet_b2b_in_store.loc[final_sheet_b2b_in_store['store_id'] == 10, 'primary_division'] = "Thokar Niaz Baig"
final_sheet_b2b_in_store.loc[final_sheet_b2b_in_store['store_id'] == 11, 'primary_division'] = "Capital"
final_sheet_b2b_in_store.loc[final_sheet_b2b_in_store['store_id'] == 12, 'primary_division'] = "Safari"
final_sheet_b2b_in_store.loc[final_sheet_b2b_in_store['store_id'] == 15, 'primary_division'] = "Faisalabad"
final_sheet_b2b_in_store.loc[final_sheet_b2b_in_store['store_id'] == 16, 'primary_division'] = "DHA Store"
final_sheet_b2b_in_store.loc[final_sheet_b2b_in_store['store_id'] == 22, 'primary_division'] = "Manghopir"
final_sheet_b2b_in_store.loc[final_sheet_b2b_in_store['store_id'] == 23, 'primary_division'] = "Star Gate"
final_sheet_b2b_in_store.loc[final_sheet_b2b_in_store['store_id'] == 25, 'primary_division'] = "Ravi Road"
final_sheet_b2b_in_store.loc[final_sheet_b2b_in_store['store_id'] == 26, 'primary_division'] = "Model Town"

#adding column for medium
final_sheet_b2b_in_store['medium'] = 'SMS'

#adding column for feedback group type
final_sheet_b2b_in_store['type'] = 'B2B In-Store'

#converting segment from O&I to Office and Industry
final_sheet_b2b_in_store['segment'].loc[(final_sheet_b2b_in_store['segment'] == 'O&I')] = "Office and Industry"

#standardizing column names for final sheet B2B In-Store
final_sheet_b2b_in_store.columns = final_sheet_column_names

#exporting to CSV
final_sheet_b2b_in_store.to_csv("B2B In-Store.csv", index=False)

#################################################################
################Sending Out the Email############################
#################################################################

#listing down email address for senders and recepients
sender_email = 'adi.gillani@gmail.com'
recipient_email = ['adi.gillani@gmail.com','adnan.gillani@sentimeter.io']

#composing message for email subject and body
fsd_email_subject = 'B2B FSD' + ' ' + (datetime.now() - timedelta(1)).strftime('%d-%m-%Y')
fsd_email_body = 'B2B FSD data for ' + (datetime.now() - timedelta(1)).strftime('%d-%m-%Y')

b2b_in_store_subject = 'B2B In-Store' + ' ' + (datetime.now() - timedelta(1)).strftime('%d-%m-%Y')
b2b_in_store_body = 'B2B In-Store data for ' + (datetime.now() - timedelta(1)).strftime('%d-%m-%Y')

#initializing server connection to send email
yag = yagmail.SMTP(user='adi.gillani@gmail.com', password='helterskelter')

#sending out email for B2B FSD
yag.send(
    to=recipient_email,
    subject=fsd_email_subject,
    contents=fsd_email_body,
    attachments='B2B FSD.csv'
)

#sending out email for B2B In-Store
yag.send(
    to=recipient_email,
    subject=b2b_in_store_subject,
    contents=b2b_in_store_body,
    attachments='B2B In-Store.csv'
)