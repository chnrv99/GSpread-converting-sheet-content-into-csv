from flask import Flask, render_template, request
import gspread
import csv

app = Flask(__name__)

def convertToCSV(link, mainSheetName, sheetName):
    email = 'gspread-account-84@forward-quanta-390017.iam.gserviceaccount.com'


    print('Email is:', email, ' Add the email to the sharing link')

    sa = gspread.service_account(filename='credentials.json')

    # mainSheetName = 'Email Subscribers'
    sh = sa.open(mainSheetName)


    # sheetName = 'Sheet1'
    wks = sh.worksheet(sheetName)

    print('Rows: ', wks.row_count)
    print('Column: ', wks.col_count)

    # print(wks.acell('A1').value)
    # print(wks.cell(2,2).value)

    # print(wks.get('A1:B2'))
    # dict
    print(wks.get_all_records())
    # list
    # print(wks.get_all_values())


    my_dict = wks.get_all_records()
    headings = ['Name', 'Email']

    with open('csvfile.csv', 'w') as f:
        # w = csv.DictWriter(f, my_dict.keys())
        w = csv.DictWriter(f, fieldnames=headings)
        w.writeheader()
        w.writerows(my_dict)
    return 'csvfile.csv'

@app.route('/api/upload',methods=['POST'])
def convert():
    if request.method == 'POST':
        link = request.form.get('Link')
        mainSheetName = request.form.get('mainSheetName')
        sheetName = request.form.get('sheetName')
        name = convertToCSV(link, mainSheetName, sheetName)
        return name
