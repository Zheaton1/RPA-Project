#question 1
while True:
    directory = single_input('Please input the directory to the folder with the following: `ProductData.xlsx`, `MasterInvoice.pdf`, `Order1.xlsx`, `Order2.xlsx`, `Order3.xlsx`, and `Order4.xlsx`')
    if directory.endswith('/') == False:
        directory = directory + '/'

    try:
        filenames = os.listdir(directory)
        if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
            break
        else:
            show_message('Error','Invalid Input: please re-enter the directory')
    except:
        show_message('Error','Invalid Input: please re-enter the directory')
       

# '/Users/zane/Documents/School/ACC 525/17 - Project 3 - Robotic Process Automation'

#question 2
product_descriptions = {}
product_prices = {}

df = pd.read_excel(f'{directory}ProductData.xlsx', sheet_name='Descriptions')
for index, row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df2 = pd.read_excel(f'{directory}ProductData.xlsx', sheet_name='Prices')
for index, row in df2.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

print(product_descriptions)
print(product_prices)

#question 3
def get_order_data(invoice_num,filename,directory):
    workbook = openpyxl.load_workbook(f'{directory}{filename}', data_only=True)
    worksheet = workbook['OrderForm']

    # Given information
    name = worksheet['C14'].value
    email = worksheet['C13'].value
    date = worksheet['C9'].value
    date = date.strftime('%m/%d/%Y')
    order_data = {} 

    
    # Info from OrderForm
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = date
    order_data['CustomerAddress1'] = str(worksheet['C16'].value)
    order_data['CustomerName'] = str(name)
    order_data['CustomerCompany'] = str(worksheet['C15'].value)
    order_data['CustomerAddress2'] = str(worksheet['C17'].value) + ', ' + str(worksheet['C18'].value) + " " + str(worksheet['C19'].value)

    # Info for first row
    unit_number0 = str(worksheet['B24'].value)
    price0 = product_prices[unit_number0]
    order_data['Price0'] = f'${price0:,.2f}'
    order_data['Quantity0'] = str(worksheet['C24'].value)
    amount0 = float(worksheet['C24'].value) * float(product_prices[unit_number0])
    order_data['Amount0'] = f'${amount0:,.2f}'
    order_data['Description0'] = str(product_descriptions[unit_number0])
    total = amount0

    # Info if 2nd row
    if pd.isnull(worksheet['B25'].value) == False:
        unit_number1 = str(worksheet['B25'].value)
        price1 = product_prices[unit_number1]
        order_data['Price1'] = f'${price1:,.2f}'
        order_data['Quantity1'] = str(worksheet['C25'].value)
        amount1 = float(worksheet['C25'].value) * float(product_prices[unit_number1])
        order_data['Amount1'] = f'${amount1:,.2f}'
        order_data['Description1'] = str(product_descriptions[unit_number1])
        total = total + amount1

    # Info if 3rd row
    elif pd.isnull(worksheet['B27'].value) == False:
        unit_number2 = str(worksheet['B26'].value)
        price2 = product_prices[unit_number2]
        order_data['Price2'] = f'${price2:,.2f}'
        order_data['Quantity2'] = str(worksheet['C26'].value)
        amount2 = float(worksheet['C26'].value) * float(product_prices[unit_number2])
        order_data['Amount2'] = f'${amount2:,.2f}'
        order_data['Description2'] = str(product_descriptions[unit_number2])
        total = total + amount2

    # Info if 4rd row
    elif pd.isnull(worksheet['B28'].value) == False:
        unit_number3 = str(worksheet['B27'].value)
        price3 = product_prices[unit_number3]
        order_data['Price3'] = f'${price3:,.2f}'
        order_data['Quantity3'] = str(worksheet['C27'].value)
        amount3 = float(worksheet['C27'].value) * float(product_prices[unit_number3])
        order_data['Amount3'] = f'${amount3:,.2f}'
        order_data['Description3'] = str(product_descriptions[unit_number3])
        total = total + amount3

    # Info if 5th row
    elif pd.isnull(worksheet['B29'].value) == False:
        unit_number4 = str(worksheet['B28'].value)
        price4 = product_prices[unit_number4]
        order_data['Price4'] = f'${price4:,.2f}'
        order_data['Quantity4'] = str(worksheet['C28'].value)
        amount4 = float(worksheet['C28'].value) * float(product_prices[unit_number4])
        order_data['Amount4'] = f'${amount4:,.2f}'
        order_data['Description4'] = str(product_descriptions[unit_number4])
        total = total + amount4
        
    order_data['Total'] = f'${total:,.2f}'

    # order_data['Description3'] = ''
    # order_data['Description2'] = ''
    # order_data['Description1'] = ''
    # order_data['Description4'] = ''

    return order_data, name, email, date

order_data, name, email, date = get_order_data(8515,'Order1.xlsx',directory)
print(order_data,name,email,date)


#question 4

def create_invoice(invoice_num,order_data,directory):
    create_pdf(f'{directory}MasterInvoice.pdf', f'{directory}Invoice {invoice_num}.pdf', order_data)


Q4_test = create_invoice(8515,order_data,directory)
print(Q4_test)

#question 5
import smtplib
import mimetypes
from email.message import EmailMessage

def send_email(invoice_num,name,email,date,directory):

    gmail_user = "Zaneheaton@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')

    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},

Attached is your invoice dated {date}

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment
    path_to_file = f'{directory}Invoice {invoice_num}.pdf' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,'Thomas Holt','zaneheaton@gmail.com','11/13/2023',directory)

#question 6
filenames = os.listdir(directory)
filenames.sort()
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(invoice_num,file,directory)
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num,name,email,date,directory)
        invoice_num = invoice_num + 1
