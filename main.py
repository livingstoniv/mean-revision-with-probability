from datetime import datetime
from yahoo_fin import stock_info as si
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell
import openpyxl as px
import time


sp_tickers = si.tickers_sp500()

dow_data = {}

today = datetime.today()
strf_today = today.strftime("%Y-%m-%d %H-%M-%S")


for ticker in sp_tickers:
    #Get histprical data for stock ticker
    data = si.get_data(ticker)
    #Calculate the mean price for last 30 days
    mean_price = data['close'][-21:].mean()
    #Calculate the probability of percentage of stock price going below or above the mean
    below_mean_prob = round((data["close"] < mean_price).sum() / len(data["close"]) * 100, 2)
    above_mean_prob = round((data['close'] > mean_price).sum() / len(data['close']) * 100, 2)
    round_below_mean_prob = round(below_mean_prob, 2)
    round_above_mean_prob = round(above_mean_prob, 2)
    live_price = si.get_live_price(ticker)

    dow_data[ticker] = {'Ticker': ticker, 'Live Price': live_price, 'Mean Price': mean_price, 'Below Mean Probability': below_mean_prob, 'Above Mean Probability': above_mean_prob}
    print("Finished Ticker: " + ticker)
dow_df = pd.DataFrame.from_dict(dow_data, orient="index")
print(dow_df)


file_name = "S&P500" + "-" + str(strf_today)
sheet_name_var = "S&P500"
file_extension = '.xlsx'



    # #Code to Make Titles Of Each Column in Excel Sheet Fit in The Excel Sheet Cells By Iterating Through Each Column And Resizing It To Fit Titles
writer = pd.ExcelWriter(file_name + file_extension, engine='xlsxwriter')
dow_df.to_excel(writer, sheet_name=sheet_name_var, index=False, na_rep='NaN')
for column in dow_df:
    column_length = max(dow_df[column].astype(str).map(len).max(), len(column))
    col_idx = dow_df.columns.get_loc(column)
    writer.sheets[sheet_name_var].set_column(col_idx, col_idx, column_length)
#Saves The Excel Sheet To The Folder That Holds Your Python Files
writer.save() 
