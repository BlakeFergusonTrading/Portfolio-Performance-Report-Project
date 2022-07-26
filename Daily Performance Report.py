import pandas as pd
import numpy as np
import datetime as dt
import matplotlib.pyplot as plt
import finnhub
import json
import pandas_market_calendars as mcal
import os
import base64
from datetime import timedelta
from pathlib import Path
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Mail, Attachment, FileContent, FileName, FileType, Disposition)
from apscheduler.schedulers.background import BlockingScheduler
from statistics import mean

def create_report():

    excel_data = pd.read_excel(r'C:\Users\Blake Ferguson\OneDrive\Documents\portfolios_input.xlsx')
    port_data = pd.DataFrame(excel_data)

    cli = finnhub.Client(api_key='FinnhubAPIKey')
    quotes = []
    for symbol in excel_data['Symbol']:
        quotes.append(cli.quote([symbol]))
        performance_data = pd.DataFrame(quotes)

    performance_data.drop(columns=['d','h','l','pc','t'],inplace=True)
    performance_data['Shares'] = excel_data['Shares']

    position_entry = pd.Series(excel_data['Entry Price'] * excel_data['Shares'])
    position_value = pd.Series(performance_data['c'] * excel_data['Shares'])
    pnl = pd.Series(position_value - position_entry)

    performance_data['Entry Price'] = position_entry
    performance_data['Position Value'] = position_value
    performance_data['Profit/Loss'] = pnl

    sentiment_avg = []

    now = dt.datetime.now()
    yesterday = now - dt.timedelta(days=1)

    for symbol in excel_data['Symbol']:
        social_data = cli.stock_social_sentiment([symbol], yesterday, now)

        scores = []
        for data in social_data['twitter']:
            if data['score'] != 0:
                scores.append(data['score'])

        if len(scores) > 0:
            sentiment_avg.append(mean(scores))
        else:
            sentiment_avg.append(0)

    performance_data['Twitter Sentiment'] = sentiment_avg
    performance_data.set_index(excel_data['Symbol'],inplace=True)

    analyst_rec = {}
    for symbol in excel_data['Symbol']:
        analyst_rec = (cli.recommendation_trends([symbol]))

    highest_value = 0
    highest_rating = None
    for key, value in analyst_rec[0].items():
        if type(value) is int:
            if value > highest_value:
                highest_rating = key
                highest_value = value

    performance_data['Weighted Analyst Trend'] = highest_rating
    performance_data.columns = ['Close', '% Change', 'Open', 'Shares', 'Entry Price', 'Position Value', 'Profit/Loss', 'Twitter Sentiment', 'Weighted Analyst Trend']

    performance_data = performance_data.reindex(columns=['Profit/Loss', 'Position Value', 'Entry Price', 'Shares', 'Open', 'Close', '% Change', 'Twitter Sentiment', 'Weighted Analyst Trend'])
    performance_data.to_csv('Performance_Report.csv')

def send_report():

    FILE_PATH = "Performance_Report.csv"
    today = dt.datetime.today()

    message = Mail(
        from_email = 'example@email.com',
        to_emails = 'example@email.com',
        subject = f'Portfolio Report: {today:%B %d, %Y}',
        html_content = '<strong>Portfolio Report attached </strong>'
    )

    with open(FILE_PATH, 'rb') as f:
        data = f.read()
        f.close()

    encoded_file = base64.b64encode(data).decode()

    attachedFile = Attachment(
        FileContent(encoded_file),
        FileName(f'{Path(FILE_PATH).stem}.csv'),
        FileType('application/csv'),
        Disposition('attachment')
    )
    message.attachment = attachedFile

    sg = SendGridAPIClient('SendGridAPIKey')
    response = sg.send(message)
    if not str(response.status_code).startswith('2'):
        raise Exception(f'Could not send email: ({response.status_code})')

def email_performance_report():
    create_report()
    send_report()
    
if __name__ == '__main__':
    
    nyse = mcal.get_calendar('NYSE')
    days = nyse.schedule(start_date='2022-07-22', end_date='2027-07-22')

    schedule = BlockingScheduler()

    for date in days['market_close']:
        schedule.add_job(email_performance_report, 'date', run_date=date.strftime('%Y-%m-%d'))

    schedule.start()