from flask import Flask, render_template, redirect, url_for, flash, abort, request, send_from_directory
from flask_bootstrap import Bootstrap
from pathlib import Path
import re
from datetime import date
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.orm import relationship
from flask_login import UserMixin, login_user, LoginManager, login_required, current_user, logout_user
from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField, PasswordField, FileField
from wtforms.validators import DataRequired, URL
from flask_gravatar import Gravatar
import pdfplumber
import pandas as pd
from collections import namedtuple
from openpyxl import Workbook, load_workbook

import xlsxwriter
import csv

global upi_match_length, upi_unmatch_length, card_match_length, card_unmatch_length
import os


user = os.path.expanduser("~");
directory  =  user + '/Documents/recon_junks'
#os.makedirs(directory)

app = Flask(__name__)
app.config['SECRET_KEY'] = '8BYkEfBA6O6donzWlSihBXox7C0sKR6b'
app.config['UPLOAD_FOLDER'] = ''
Bootstrap(app)


login_details = {
    '1991':False,
    '1992':False,
    '1993':False,
    '1994':False,
    '1995':False,
    '1996':False,
    '1997':False,
    '1998':False,
    '1999':False
}


class reconciliationForm(FlaskForm):
    bankFile = FileField("Bank Statement", validators=[DataRequired()])
    upiFile = FileField("UPI Statement", validators=[DataRequired()])
    cardFile = FileField("Card Statement", validators=[DataRequired()])
    submit = SubmitField("Reconcilation")

class loginForm(FlaskForm):
    pin = StringField("Enter Your Pin", validators=[DataRequired()])
    submit = SubmitField("Submit")

def make_directory():
    global paths
    if not os.path.exists(paths):
        os.makedirs(paths)
        return

def loading_files():  # loading all the files required
    global file, df_upipine, df_cardpine, settle_upi
    file = (filename1)  # copying the pdf bank statment
    df_upipine = pd.read_excel(filename2).astype(str)  # copying the card pine statement and converting it to s str
    card_pine_exceel = pd.read_excel(filename3)
    card_pine = card_pine_exceel.loc[card_pine_exceel['Me Code'] == 'TD1882']
    settle_upi = []

    for list in df_upipine[
        'Txn Time']:  # changing the format of Txn Time in upi pine statement to match the bank statment format
        list = str(list)
        date = list.split(' ')[0]
        settle_upi.append(date)
    df_upipine['Txn Time'] = pd.DataFrame(settle_upi)

    df_cardpine = card_pine.astype(str)  # copying the upi pine statement and converting it to a str
    return


def even_number_background(cell_value):
    highlight_match = 'background-color: #90EE90;'
    highlight_unmatch = 'background-color: red;'
    highlight_failed = 'background-color: yellow;'
    if cell_value == 'Matched':
        return highlight_match
    elif cell_value == 'Mismatched Approval Code':
        return highlight_unmatch
    elif cell_value == 'Mismatched Settled Date':
        return highlight_unmatch
    elif cell_value == 'Mismatched Transaction Amount':
        return highlight_unmatch
    elif cell_value == 'Failed':
        return highlight_failed


def converting_card():  # converting the card details from bank statement to a dataframe
    global dfcard
    months = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug',
              '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'}
    try:
        LineCard = namedtuple('Line',
                              'rec_fmt bat_nbr card_type CAT card_number trans_date settle_date approv_code gross_amount_intnl domestic mdr_rate mdr_flat')
    except:
        LineCard = namedtuple('Line',
                              'rec_fmt bat_nbr card_type CAT card_number trans_date settle_date approv_code gross_amount_intnl domestic mdr_rate')
    linescard = []
    bat_re = re.compile(r'(BAT \d)')
    # opening the pdf file and reading each line, if a line starts with 'BAT ' we split that line at every space and append it to columns of LineCard
    with pdfplumber.open(file, password="TD1882") as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            for line in text.split('\n'):
                card = bat_re.search(line)
                if card:
                    # here we are removing the decimal points after amount
                    card = str(line).replace('.00', '')
                    card = str(line).replace('(cid:13)', '')
                    card_details = card.split()
                    # here we are removing the leading zeros
                    try:
                        card_detail = [(lambda x: x.lstrip('0') if isinstance(x, str) and len(x) != 1 else x)(x) for x
                                       in card_details]
                        linescard.append(LineCard(*card_detail))
                    except TypeError:
                        loop_ender = False
                        while loop_ender == False:
                            for line_by_line in card_details:
                                x = re.search("^[*0-9]+\.[0-9][0-9][*0-9]+\.[0-9][0-9]$", line_by_line)
                                if (x):
                                    a = line_by_line
                                    rep = 0;
                                    tress = False
                                    while tress == False:
                                        if a[rep] == '.':
                                            tress = True
                                            rest = a[slice(0, rep + 3)]
                                            b = a.split(rest, 1)
                                            y = card_details.index(line_by_line)
                                            z = line_by_line.split("0.00", 1)[1]
                                            card_details[y] = rest
                                        rep += 1

                                    card_details.insert(y + 1, b[1])
                                    try:
                                        card_detail = [
                                            (lambda x: x.lstrip('0') if isinstance(x, str) and len(x) != 1 else x)(x)
                                            for x in card_details]
                                        linescard.append(LineCard(*card_detail))
                                        loop_ender = True
                                    except TypeError:
                                        loop_ender = False

    dfupi_mont = pd.DataFrame(linescard)
    dfupi_months = dfupi_mont['settle_date'][0].split('-')[1]
    month_array = []
    final_months = months[dfupi_months]
    for mon in dfupi_mont['settle_date']:
        day = dfupi_mont['settle_date'][0].split('-')[0]
        mont = final_months
        year = dfupi_mont['settle_date'][0].split('-')[2]
        date = day + '-' + mont + '-' + year
        month_array.append(date)
    dfupi_mont = pd.DataFrame(linescard)
    dfupi_mont['settle_date'] = month_array
    dfcard = dfupi_mont.astype(str)
    return dfcard


def converting_upi():  # converting upi details of the bank statement into a dataframe
    global dfupi
    LineUPI = namedtuple('Line', 'rec_fmt pay_type payer_vpa trans_date settle_date amount_domestic rrn_no UPI_txn_id')
    linesupi = []
    cr_re = re.compile(r'(CR )')
    # opening the pdf file and reading each line, if a line starts with 'CR ' we split that line at every space and append it to columns of LineCard#
    with pdfplumber.open(file, password="TD1882") as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            for line in text.split('\n'):
                upi = cr_re.search(line)
                if upi:
                    # here we are removing the decimal points after amount
                    upi = str(line).replace('.00', '')
                    upi = str(line).replace('(cid:13)', '')
                    upi_details = upi.split()
                    upi_details = [(lambda x: x.lstrip('0') if isinstance(x, str) and len(x) != 1 else x)(x) for x in
                                   upi_details]
                    linesupi.append(LineUPI(*upi_details))

    dfupi = pd.DataFrame(linesupi).astype(str)
    return dfupi


def reconciliation():  # getting all the matched and difference dataframes after comparing it
    global df_cardbank_diff, df_cardpine_diff, df_upibank_diff, df_upipine_diff, df_cardbank_match, df_cardpine_match, df_upibank_match, df_upipine_match, df_cardbank_settle_match, df_cardpine_settle_match, df_cardbank_settle_diff, df_cardpine_settle_diff, df_upibank_diffreason, df_upipine_diffreason, df_cardbank_diffreason, df_cardpine_diffreason, df_total, procesed, today_date
    dfupi['rrn_no'] = pd.to_numeric(dfupi['rrn_no'], errors='coerce')  # carder['domestic'][0]
    dfcard['domestic'] = pd.to_numeric(dfcard['domestic'], errors='coerce')  # carder['domestic'][0]
    df_upipine['RRN'] = pd.to_numeric(df_upipine['RRN'], errors='coerce')
    df_cardpine['TRANSACTION AMOUNT'] = pd.to_numeric(df_cardpine['TRANSACTION AMOUNT'], errors='coerce')
    dfupi['amount_domestic'] = pd.to_numeric(dfupi['amount_domestic'], errors='coerce')
    df_upipine['Txn Amount'] = pd.to_numeric(df_upipine['Txn Amount'], errors='coerce')
    df_upipine['Txn Time'] = pd.to_datetime(df_upipine['Txn Time'])
    dfupi['trans_date'] = pd.to_datetime(dfupi['trans_date'])

    today_date = df_cardpine['SETTLED DATE'].loc[0]
    #if (procesed):
        #df_total = dfcard.shape[0] + df_cardpine.shape[0] + df_upipine.shape[0] + dfupi.shape[0]

    dfcard['approv_code'] = dfcard['approv_code'].astype(str)
    df_cardpine['APPROVAL CODE'] = df_cardpine['APPROVAL CODE'].astype(str)

    df_upibank_diff = dfupi[~(
                (dfupi['rrn_no'].isin(df_upipine['RRN'])) & (dfupi['trans_date'].isin(df_upipine['Txn Time'])) & (
            dfupi['amount_domestic'].isin(df_upipine['Txn Amount'])))].sort_values('rrn_no').reset_index(
        drop=True).replace('nan', ' ')  # upi transactions missing from bank statement
    df_upipine_diff = df_upipine[~ (
                (df_upipine['RRN'].isin(dfupi['rrn_no'])) & (df_upipine['Txn Time'].isin(dfupi['trans_date'])) & (
            df_upipine['Txn Amount'].isin(dfupi['amount_domestic'])))].sort_values('RRN').reset_index(
        drop=True).replace('nan', ' ')  # upi transactions missing from pine statement
    df_upibank_match = dfupi[~(
                (dfupi['rrn_no'].isin(df_upipine['RRN'])) & (dfupi['trans_date'].isin(df_upipine['Txn Time'])) & (
            dfupi['amount_domestic'].isin(df_upipine['Txn Amount']))) == False].sort_values('rrn_no').reset_index(
        drop=True).replace('nan', ' ')  # df of upi transactions present in both statements (bank columns)
    df_upipine_match = df_upipine[~(
                (df_upipine['RRN'].isin(dfupi['rrn_no'])) & (df_upipine['Txn Time'].isin(dfupi['trans_date'])) & (
            df_upipine['Txn Amount'].isin(dfupi['amount_domestic']))) == False].sort_values('RRN').reset_index(
        drop=True).replace('nan', ' ')  # df of upi transactions present in both statements (pine columns)
    df_cardbank_diff = dfcard[~((dfcard['approv_code'].isin(df_cardpine['APPROVAL CODE'])) & (
        dfcard['settle_date'].isin(df_cardpine['SETTLED DATE'])) & (
                                    dfcard['domestic'].isin(df_cardpine['TRANSACTION AMOUNT'])))].sort_values(
        'approv_code').reset_index(drop=True).replace('nan', ' ')  # card transactions missing from bank statement

    df_cardbank_approv = (df_cardbank_diff[~((df_cardbank_diff['approv_code'].isin(df_cardpine['APPROVAL CODE'])))])
    df_cardbank_approv['Status'] = 'Mismatched Approval Code'
    df_cardbank_settle = (df_cardbank_diff[~((df_cardbank_diff['settle_date'].isin(df_cardpine['SETTLED DATE'])))])
    df_cardbank_settle['Status'] = 'Mismatched Settled Date'
    df_cardbank_domestic = (df_cardbank_diff[~((df_cardbank_diff['domestic'].isin(df_cardpine['TRANSACTION AMOUNT'])))])
    df_cardbank_domestic['Status'] = 'Mismatched Transaction Amount'
    df_cardbank_diffreason = pd.concat([df_cardbank_approv, df_cardbank_settle, df_cardbank_domestic], axis=0,
                                       join="inner", ignore_index=False)

    df_cardpine_diff = df_cardpine[~((df_cardpine['APPROVAL CODE'].isin(dfcard['approv_code'])) & (
        df_cardpine['SETTLED DATE'].isin(dfcard['settle_date'])) & (
                                         df_cardpine['TRANSACTION AMOUNT'].isin(dfcard['domestic'])))].sort_values(
        'APPROVAL CODE').reset_index(drop=True).replace('nan', ' ')  # card transactions missing from pine statement
    df_cardpine_approv = (df_cardpine_diff[~((df_cardpine_diff['APPROVAL CODE'].isin(dfcard['approv_code'])))])
    df_cardpine_approv['Status'] = 'Mismatched Approval Code'
    df_cardpine_settle = (df_cardpine_diff[~((df_cardpine_diff['SETTLED DATE'].isin(dfcard['settle_date'])))])
    df_cardpine_settle['Status'] = 'Mismatched Settled Date'
    df_cardpine_domestic = (df_cardpine_diff[~((df_cardpine_diff['TRANSACTION AMOUNT'].isin(dfcard['domestic'])))])
    df_cardpine_domestic['Status'] = 'Mismatched Transaction Amount'
    df_cardpine_diffreason = pd.concat([df_cardpine_approv, df_cardpine_settle, df_cardpine_domestic], axis=0,
                                       join="inner", ignore_index=False)

    df_upibank_diff = dfupi[~(
                (dfupi['rrn_no'].isin(df_upipine['RRN'])) & (dfupi['trans_date'].isin(df_upipine['Txn Time'])) & (
            dfupi['amount_domestic'].isin(df_upipine['Txn Amount'])))].sort_values('rrn_no').reset_index(
        drop=True).replace('nan', ' ')  # upi transactions missing from bank statement
    df_upibank_approv = df_upibank_diff[~((df_upibank_diff['rrn_no'].isin(df_upipine['RRN'])))]
    df_upibank_approv['Status'] = 'Mismatched Approval Code'
    df_upibank_settle = df_upibank_diff[~((df_upibank_diff['trans_date'].isin(df_upipine['Txn Time'])))]
    df_upibank_settle['Status'] = 'Mismatched Settled Date'
    df_upibank_domestic = df_upibank_diff[~((df_upibank_diff['amount_domestic'].isin(df_upipine['Txn Amount'])))]
    df_upibank_domestic['Status'] = 'Mismatched Transaction Amount'
    df_upibank_diffreason = pd.concat([df_upibank_approv, df_upibank_settle, df_upibank_domestic], axis=0, join="inner",
                                      ignore_index=False)

    df_upipine_diff = df_upipine[~ (
                (df_upipine['RRN'].isin(dfupi['rrn_no'])) & (df_upipine['Txn Time'].isin(dfupi['trans_date'])) & (
            df_upipine['Txn Amount'].isin(dfupi['amount_domestic'])))].sort_values('RRN').reset_index(
        drop=True).replace('nan', ' ')  # upi transactions missing from pine statement
    df_upipine_approv = df_upipine_diff[~((df_upipine_diff['RRN'].isin(dfupi['rrn_no'])))]
    df_upipine_approv['Status'] = 'Mismatched Approval Code'
    df_upipine_settle = df_upipine_diff[~((df_upipine_diff['Txn Time'].isin(dfupi['trans_date'])))]
    df_upipine_settle['Status'] = 'Mismatched Settled Date'
    df_upipine_domestic = df_upipine_diff[~((df_upipine_diff['Txn Amount'].isin(dfupi['amount_domestic'])))]
    df_upipine_domestic['Status'] = 'Mismatched Transaction Amount'
    df_upipine_diffreason = pd.concat([df_upipine_approv, df_upipine_settle, df_upipine_domestic], axis=0, join="inner",
                                      ignore_index=False)

    df_cardbank_match = dfcard[~((dfcard['approv_code'].isin(df_cardpine['APPROVAL CODE'])) & (
        dfcard['settle_date'].isin(df_cardpine['SETTLED DATE'])) & (
                                     dfcard['domestic'].isin(df_cardpine['TRANSACTION AMOUNT']))) == False].sort_values(
        'approv_code').reset_index(drop=True).replace('nan',
                                                      ' ')  # card transactions present in both statements (bank columns)
    df_cardbank_match['Status'] = 'Matched'
    df_cardpine_match = df_cardpine[~((df_cardpine['APPROVAL CODE'].isin(dfcard['approv_code'])) & (
        df_cardpine['SETTLED DATE'].isin(dfcard['settle_date'])) & (df_cardpine['TRANSACTION AMOUNT'].isin(
        dfcard['domestic']))) == False].sort_values('APPROVAL CODE').reset_index(drop=True).replace('nan',
                                                                                                    ' ')  # card transactions present in both statements (pine columns)
    df_upibank_match = dfupi[~(
                (dfupi['rrn_no'].isin(df_upipine['RRN'])) & (dfupi['trans_date'].isin(df_upipine['Txn Time'])) & (
            dfupi['amount_domestic'].isin(df_upipine['Txn Amount']))) == False].sort_values('rrn_no').reset_index(
        drop=True).replace('nan', ' ')  # df of upi transactions present in both statements (bank columns)
    df_upibank_match['Status'] = 'Matched'
    df_upipine_match = df_upipine[~(
                (df_upipine['RRN'].isin(dfupi['rrn_no'])) & (df_upipine['Txn Time'].isin(dfupi['trans_date'])) & (
            df_upipine['Txn Amount'].isin(dfupi['amount_domestic']))) == False].sort_values('RRN').reset_index(
        drop=True).replace('nan', ' ')  # df of upi transactions present in both statements (pine columns)
    return


def remove_file():
    os.remove('Summary.xlsx')
    os.remove('summary_card.csv')
    os.remove('summary_upi.csv')
    return

def clear_history():
    global paths
    os.remove(paths + '/card_matching.csv')
    os.remove(paths + '/upi_matching.csv')
    os.remove(paths + '/card_unmatching.csv')
    os.remove(paths + '/upi_unmatching.csv')
    os.remove(paths + '/FULL_SUMMARY.csv')
    os.remove(paths + '/Reconcilation Summary.xlsx')
    os.remove(paths + '/recon_stats.csv')


def delete_failed():
    global upi_unmatching5, card_unmatching5, paths
    index_names_upi = upi_unmatching5[(upi_unmatching5['Status'] == 'Failed')].index

    index_names_card = card_unmatching5[(card_unmatching5['Status'] == 'Failed')].index
    upi_unmatching5.drop(index_names_upi, inplace=True)
    card_unmatching5.drop(index_names_card, inplace=True)
    card_unmatching5.to_csv(paths + '/card_unmatching.csv', index=False)  # storing in junk files for later use
    upi_unmatching5.to_csv(paths + '/upi_unmatching.csv', index=False)
    print('HI')
    return None

def report_download():
    global paths, upi_match_length, upi_unmatch_length, card_match_length, card_unmatch_length, failed_transaction_length, upi_unmatching5, card_unmatching5, df_totalwhole, downloads
    upi_diff = pd.concat([df_upibank_diffreason, df_upipine_diffreason], axis=0, join="outer",
                         ignore_index=False)  # concating upi diff
    card_diff = pd.concat([df_cardbank_diffreason, df_cardpine_diffreason], axis=0, join="outer",
                          ignore_index=False)  # concating card diff
    upi_Match = pd.concat([df_upibank_match, df_upipine_match], axis=1, join="outer", ignore_index=False,
                          sort=False)  # concating upi match
    card_Match = pd.concat([df_cardbank_match, df_cardpine_match], axis=1, join="outer", ignore_index=False,
                           sort=False)  # concating card diff
    empty_rrn = upi_diff.loc[upi_diff['Txn Status'] == 'TIMED OUT FAILED-NO RESPONSE FROM BANK'].drop_duplicates(
        subset=['RRN', 'rrn_no', 'Txn Id Prefix'], keep='first')
    empty_rrn['Status'] = 'Failed'

    failed_transaction_length = empty_rrn.shape[0]

    upi_differ = upi_diff.loc[upi_diff['Txn Status'] != 'TIMED OUT FAILED-NO RESPONSE FROM BANK'].drop_duplicates(
        subset=['RRN', 'rrn_no', 'Txn Id Prefix'], keep='first')

    upi_diff = pd.concat([upi_differ, empty_rrn], axis=0, join="outer", ignore_index=True)

    df_summary_card = pd.concat([card_Match, card_diff], axis=0, join="outer", ignore_index=True).drop_duplicates(
        subset=['approv_code', 'APPROVAL CODE', 'CARD NO'], keep='first')  # concaing matched and unmatched of card
    df_summary_upi = pd.concat([upi_Match, upi_diff], axis=0, join="outer", ignore_index=True).drop_duplicates(
        subset=['RRN', 'rrn_no', 'Txn Id Prefix'], keep='first')  # concaing matched and unmatched of upi

    status_column1 = card_diff.pop('Status')
    status_column2 = card_Match.pop('Status')
    status_column3 = upi_Match.pop('Status')
    status_column4 = upi_diff.pop('Status')

    # insert column using insert(position,column_name,
    # first_column) function
    card_diff.insert(12, 'Status', status_column1)
    card_Match.insert(12, 'Status', status_column2)
    upi_Match.insert(8, 'Status', status_column3)
    upi_diff.insert(8, 'Status', status_column4)

    card_matching5 = df_summary_card.loc[df_summary_card['Status'] == 'Matched']  # to segregate matched card
    card_match_length = card_matching5.shape[0]
    card_unmatching5 = df_summary_card.loc[df_summary_card['Status'] != 'Matched']  # to segregate unmatched card
    card_unmatch_length = card_unmatching5.shape[0]

    upi_matching5 = df_summary_upi.loc[df_summary_upi['Status'] == 'Matched']  # to segregate matched upi
    upi_unmatching5 = df_summary_upi.loc[df_summary_upi['Status'] != 'Matched']  # to segregate unmatched upi
    upi_match_length = upi_matching5.shape[0]
    upi_unmatch_length = upi_unmatching5.shape[0] - empty_rrn.shape[0]
    #indeed = [0, 1, 2, 3, 4, 5]
    #df_totalwhole = card_unmatch_length + card_match_length + upi_match_length + upi_unmatch_length + failed_transaction_length
    #data1 = {'Date': [today_date],
            # 'Total Transactions': [df_totalwhole],
            # 'Card Transactions Matched': [card_match_length],
            # 'Card Transactions Mismatched': [card_unmatch_length],
            # 'UPI Transactions Matched': [upi_match_length],
            # 'UPI Transactions Mismatched': [upi_unmatch_length],
            # 'Failed Transactions': [failed_transaction_length]
             #}
    #data2 = {'Date': today_date,
            # 'Total Transactions': df_totalwhole,
            # 'Card Transactions Matched': card_match_length,
            # 'Card Transactions Mismatched': card_unmatch_length,
            # 'UPI Transactions Matched': upi_match_length,
            # 'UPI Transactions Mismatched': upi_unmatch_length,
            # 'Failed Transactions': failed_transaction_length
            # }

    #if os.path.exists(directory + '/recon_stats.csv'):
        #f = open(directory + '/recon_stats.csv', 'a')
        #w = csv.DictWriter(f, data2.keys())
        #w.writerow(data2)
        #f.close()
    #else:
        #recon_stats = pd.DataFrame(data1)
        #recon_stats.to_csv(directory + '/recon_stats.csv', index=False)

    card_unmatching5.to_csv(paths + '/card_unmatching.csv', index=False)  # storing in junk files for later use
    upi_matching5.to_csv(paths + '/upi_matching.csv', index=False)
    upi_unmatching5.to_csv(paths + '/upi_unmatching.csv', index=False)
    card_matching5.to_csv(paths + '/card_matching.csv', index=False)

    df_summary_card.to_csv("summary_card.csv", index=False)
    df_summary_upi.to_csv('summary_upi.csv', index=False)
    f1 = open("summary_card.csv")
    f1_contents = f1.read()
    f1.close()

    f2 = open("summary_upi.csv")

    f2_contents = f2.read()
    f2.close()

    f3 = open(paths + '/FULL_SUMMARY.csv', 'a')  # open in `a` mode to amend
    f3.write('Filename:' + "\n" + "Bank Statement: "  + "\n" + "Pine UPI Statement: "  + "\n" + "Pine Card Statement: " + "\n" + "\n" + "\n" + 'CARD Reconciliation' + "\n" + f1_contents + "\n" + "\n" + 'UPI Reconciliation' + "\n" + f2_contents + "\n")  # concatenate the contents
    f3.close()

    wb = Workbook()
    ws = wb.active
    with open(paths + '/FULL_SUMMARY.csv', 'r') as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save('Summary.xlsx')

    excel = pd.read_excel('Summary.xlsx', header=None)
    excel2 = excel.style.applymap(even_number_background)  # coloring matched and unmatched
    excel2.to_excel(paths + '/Reconcilation Summary.xlsx', index=False)  # final output
    #downloads = excel2.to_excel('Reconcilation Summary.xlsx', index=False)
    remove_file()  # to remove unwanted file
    return None


def segregating(df_existing_upi_unmatch, card_upi):
    global unmatched_card_exsisting, unmatched_upi_exsisting, unmatched_card_bank_exsisting, unmatched_card_pine_exsisting, unmatched_upi_pine_exsisting, unmatched_upi_bank_exsisting
    unmatched_card_exsisting = pd.DataFrame()
    unmatched_bank_exsisting_count = 0
    unmatched_pine_exsisting_count = 0
    unmatched_upi_exsisting = pd.DataFrame()
    if (card_upi == 'card'):  # contains starting and ending column of card details
        row_para = 'rec_fmt'
        row_value = 'BAT'
        row1_start = 'rec_fmt'
        row2_start = 'Me Code'
        row1_end = 'mdr_flat'
        row2_end = 'GSTN ID'
    else:  # contains starting and ending column of upi details
        row_para = 'rec_fmt'
        row_value = 'CR'
        row1_start = 'rec_fmt'
        row2_start = 'Sr. No.'
        row1_end = 'UPI_txn_id'
        row2_end = 'Source'
    card_unmatch_len = df_existing_upi_unmatch.shape[0]  # calculating row of respective
    if (card_unmatch_len == 0):  # if no of row is zero
        unmatched_card_exsisting = pd.DataFrame(df_existing_upi_unmatch.loc[:, row1_start:row1_end])
        unmatched_upi_exsisting = pd.DataFrame(df_existing_upi_unmatch.loc[:, row2_start:row2_end])
    else:
        for unmatch_position in range(card_unmatch_len):
            approve_code = df_existing_upi_unmatch.loc[[unmatch_position]]  # position of each row

            bank_or_upi = (approve_code[row_para] == row_value).any()  # if first line starting with card/upi or pine
            if (bank_or_upi):
                if (unmatched_bank_exsisting_count == 0):
                    unmatched_card_exsisting = pd.DataFrame(df_existing_upi_unmatch.loc[[unmatch_position]].loc[:,
                                                            row1_start:row1_end])  # sending data to datafame of existing
                    unmatched_bank_exsisting_count += 1
                else:
                    unmatched_card_exsisting = pd.concat([unmatched_card_exsisting,
                                                          df_existing_upi_unmatch.loc[[unmatch_position]].loc[:,
                                                          row1_start:row1_end]], axis=0, join="outer",
                                                         ignore_index=False)  # appending data to datafame of existing


            else:
                if (unmatched_pine_exsisting_count == 0):
                    unmatched_upi_exsisting = pd.DataFrame(df_existing_upi_unmatch.loc[[unmatch_position]].loc[:,
                                                           row2_start:row2_end])  # sending data to datafame of existing
                    unmatched_pine_exsisting_count += 1
                else:
                    unmatched_upi_exsisting = pd.concat([unmatched_upi_exsisting,
                                                         df_existing_upi_unmatch.loc[[unmatch_position]].loc[:,
                                                         row2_start:row2_end]], axis=0, join="outer",
                                                        ignore_index=False)  # appending data to datafame of existing

    if (card_upi == 'card'):  # assaigning to cards/upi
        unmatched_card_pine_exsisting = unmatched_upi_exsisting
        unmatched_card_bank_exsisting = unmatched_card_exsisting
        return

    else:
        unmatched_upi_pine_exsisting = unmatched_upi_exsisting
        unmatched_upi_bank_exsisting = unmatched_card_exsisting
        return
def file_creation(cofe):
    global paths
    paths = str(Path().absolute()) + '/' + cofe

def existing_file_concatenation():
    global dfcard,  df_cardpine, dfupi, df_upipine,df_total
    df_existing_card_unmatch = pd.read_csv(paths + '/card_unmatching.csv')#reading unmatched card detaiils
    df_existing_upi_unmatch = pd.read_csv(paths + '/upi_unmatching.csv')#reading unmatched upi detaiils
    segregating(df_existing_card_unmatch,'card')
    segregating(df_existing_upi_unmatch,'upi')
    df_total = dfcard.shape[0] + df_cardpine.shape[0] + df_upipine.shape[0] + dfupi.shape[0]

    dfcard = pd.concat([dfcard, unmatched_card_bank_exsisting], axis=0, join="outer",ignore_index=True, verify_integrity=False).drop_duplicates(subset=['approv_code', 'card_number'], keep='first')#concating bank card details with existing
    df_cardpine = pd.concat([df_cardpine, unmatched_card_pine_exsisting], axis=0, join="outer",ignore_index=True, verify_integrity=False).drop_duplicates(subset=['APPROVAL CODE', 'CARD NO'], keep='first')#concating pine card with existing
    df_upipine = pd.concat([df_upipine, unmatched_upi_pine_exsisting], axis=0, join="outer",ignore_index=True, verify_integrity=False).drop_duplicates(subset=['Host Txn Id'], keep='first')#concating upi pine with existing
    dfupi = pd.concat([dfupi, unmatched_upi_bank_exsisting], axis=0, join="outer",ignore_index=True, verify_integrity=False).drop_duplicates(subset=['rrn_no'], keep='first')#concating card upi with existing

def summary_check():
    global procesed,paths
    openFile = True
    try:
        file = pd.read_excel(paths + '/Reconcilation Summary.xlsx', header=None)  # checking if this file exsist or not
    except FileNotFoundError:
        openFile = False  # if file is not there
    if (openFile == True):
        print('hai')
        procesed = False
        existing_file_concatenation()
        reconciliation()
        report_download()

    else:
        procesed = True
        make_directory()
        reconciliation()
        report_download()

    return
@app.route("/logout", methods=["GET", "POST"])
def logout():
    global code
    login_details[code] = False
    return redirect(url_for('login'))
@app.route("/", methods=["GET", "POST"])
def login():
    global code
    login = loginForm()
    if login.validate_on_submit():
        code = login.pin.data;
        print(login_details[code])
        if(login_details[code] == False):
            login_details[code] = True
            file_creation(code);
            return redirect(url_for('hello_world'))
        else:
            return render_template('error page.html')
    return render_template('login page.html', forms=login)
@app.route("/recon", methods=["GET", "POST"])
def hello_world():
    global filename1, filename2, filename3
    file_creation(code);
    forms = reconciliationForm()
    if forms.validate_on_submit():
        print(forms.bankFile.data)
        print(forms.upiFile.data)
        print(forms.cardFile.data)
        filename1 = forms.bankFile.data;
        filename2 = forms.upiFile.data;
        filename3 = forms.cardFile.data;
        loading_files()
        converting_card()
        converting_upi()
        return redirect(url_for('flask_reconcilation'))
    return render_template('recon.html', forms=forms)

@app.route("/reconciling",methods=["GET","POST"])
def flask_reconcilation():
    summary_check()

    return redirect(url_for('hello_world'))
@app.route("/download",methods=["GET","POST"])
def download_file():
    global paths
    try:
        file = pd.read_excel(paths + '/Reconcilation Summary.xlsx', header=None)  # checking if this file exsist or not'
        full_path = paths
        print(full_path)
        return send_from_directory(full_path,  'Reconcilation Summary.xlsx')
    except FileNotFoundError:
        openFile = False


    return redirect(url_for('hello_world'))

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=3000)