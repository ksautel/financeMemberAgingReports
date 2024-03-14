import pandas as pd
from docxtpl import DocxTemplate
import datetime

# User date input:
date = str(input(
    "Please enter the date exactly as you would like it to appear on the letter. "))

# User prior letter date input: 
previous_date = str(input(
    "For 60 day letters, please enter the date of the previously sent 30-Day Letter exactly as you would like it to appear on the letter."))

t1 = datetime.datetime.now()

# Document template
tmp = DocxTemplate('./60_Day_Letter.docx')

############# Function Definition #############
# File import function


# def imp(filePath):
df = pd.read_csv(r'./MemberAgingRpt.csv', header=0,
                 converters={'PastDue3': lambda s: float(s.replace('$', ''))})
# return df

# Document generation function


def document(context, fileName):
    tmp.render(context)
    tmp.save(fileName) 
    return


############# Data Selection & Document Generation Loop #############
# Extracting names and past due amounts from data frame when pastDue1 is greater than 0
first_names = df.loc[df['PastDue3'] > 0, "MemberFirstName"]
last_names = df.loc[df['PastDue3'] > 0, "MemberLastName"]
amounts = df.loc[df['PastDue3'] > 0, "PastDue3"]
emails = df.loc[df['PastDue3'] > 0, "MemberEmailAddress"]


# Looping over the list to create the documents
for i in range(len(first_names)):
    first = first_names.iloc[i]
    last = last_names.iloc[i]
    amount = amounts.iloc[i]
    context = {
        'date': date,
        'date2': previous_date,
        'first': first,
        'last': last,
        'amount': amount
    }
    documentName = str(f"./Letters/{last} 60 Day Letter")
    document(context, documentName)
    i = i + 1
    # end for

t2 = datetime.datetime.now()
delta = t2 - t1
print("This script created ", i, " documents in ", delta, "seconds.")
