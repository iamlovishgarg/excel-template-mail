from pandas import read_excel
from halo import Halo
from time import sleep
import filenames
from email.message import EmailMessage
import ssl
import smtplib
from credentials import SENDER_EMAIL, PASSWORD

# list of reserved keywords in python
PROHIBITED_WORDS = [
    "False",
    "def",
    "if",
    "raise",
    "None",
    "del",
    "import",
    "return",
    "True",
    "elif",
    "in",
    "try",
    "and",
    "else",
    "is",
    "while",
    "as",
    "except",
    "lambda",
    "with",
    "assert",
    "finally",
    "nonlocal",
    "yield",
    "break",
    "for",
    "not",
    "class",
    "form",
    "or",
    "continue",
    "global",
    "pass"
]

# function to send email
def send_email(name, receiver_email, message):
    spinner = Halo(text=f"Sending email to {name}...", spinner="dots")
    spinner.start()

    whole_email = EmailMessage()
    whole_email['From'] = SENDER_EMAIL
    whole_email['To'] = receiver_email
    whole_email['Subject'] = f"This is an automated email for {name}"
    whole_email.set_content(message)

    smtp.sendmail(SENDER_EMAIL, receiver_email, whole_email.as_string())

    spinner.succeed(f"Email sent to {name}!")

# opening template file
with open(filenames.TEMPLATE_FILENAME, "r") as f:
    template = f.read()

# opening excel file
data = read_excel(filenames.EXCEL_FILENAME)

# getting all column names
columns = list(data.columns.values) # it is a numpy array, type casting in list so that i can remove "emails"

# checking if there is an emails column
assert "emails" in columns, 'Error: No "emails" column found'
# removing it because we are mapping all columns afterwards
columns.remove("emails")

assert "emails" not in template, 'Error: "emails" can\'t be used in template as it is ignored while making template because "emails" column is used to send emails. Although, you can copy "emails" column and save it with different name and use that'


# CHECKING IF ALL COLUMN LENGTHS ARE EQUAL
all_column_lengths = []
for column in columns:
    # error if variables in template are not in excel or if template is different from excel
    assert column in template, "Error: Template file doesn't match excel file"

    # error if column name is a reserved keyword
    assert column not in PROHIBITED_WORDS, f"Error: {column} word can't be used, kindly change it"

    # appending column length
    all_column_lengths.append(data[column].count())

# if set is more than 1 then column length is not same as all items in the list should be same
assert len(set(all_column_lengths))==1, "Error: Column length is not same"

# CHANGING ITEMS IN "columns"
# saving first column's name before changing it to iterate later
first_column = columns[0]

# changing column name from "name" to data['name'][i] format, like this: ["data['name'][i]", "data['class_name'][i]", "data['roll_no'][i]"]
for count, column_value in enumerate(columns):
    columns[count] = f"{column_value}=data['{column_value}'][i]"
# changing list to string to directly use it, like this: "data['name'][i],data['class_name'][i],data['roll_no'][i]"
columns = ','.join(columns)

# DISPLAYING EXAMPLE
# there is "i" in "data['name'][i]", so displaying for 0th (1st row in excel)
i=0

# making query
final_query = f"final_str = '''{template}'''.format({columns})"

# executing query
exec(final_query)

# displaying final format
print("\n" + final_str)


# asking if the format is correct
ask = input("\nIs the format okay? [y/n]: ").lower()

if not (ask=="y" or ask=="yes"):
    exit() # exiting if not correct

# If everything is right then, logging in gmail and sending mails
context = ssl.create_default_context()
with smtplib.SMTP_SSL('smtp.gmail.com', port=465, context=context) as smtp:
    smtp.login(SENDER_EMAIL, PASSWORD)

    # proceeding if correct
    for i in range(len(data[first_column])):
        
        final_query = f"final_str = '''{template}'''.format({columns})"
        exec(final_query)

        # sending mail
        send_email(name=data[first_column][i], receiver_email=data['emails'][i], message=final_str)
        sleep(3)