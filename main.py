import yfinance as yf
import math
import datetime
import pandas as pd
import numpy as np
from scipy.stats import norm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time


def get_volatility_for_dates(url, date1, date2):
    driver = webdriver.Chrome()
    driver.get(url)
    time.sleep(1)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    # Set first date
    date1_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "anychart-label-input")))
    date1_input.send_keys(Keys.BACK_SPACE * 10, date1, Keys.ENTER)

    # Set second date
    driver.find_elements(By.CLASS_NAME, "anychart-label-input")[1].send_keys(Keys.BACK_SPACE * 10, date2, Keys.ENTER)

    # Extract volatility info
    soup = BeautifulSoup(driver.page_source, "html.parser")
    tspan_element = soup.find("tspan", string=lambda string: "30-Day Implied Volatility (Puts):" in string)
    implied_volatility = tspan_element.text.split(":")[1].strip()

    driver.close()

    return float(implied_volatility)


N = norm.cdf


def BS_PUT(S, K, r, T, sigma):
    d1 = (np.log(S / K) + (r + sigma ** 2 / 2) * T) / (sigma * np.sqrt(T))
    d2 = d1 - sigma * np.sqrt(T)
    return K * np.exp(-r * T) * N(-d2) - S * N(-d1)


def calculate_put_price(stock_price, implied_volatility, time_to_expiration, strike_price):
    day_after_desired_date = (
                datetime.datetime.strptime(desired_date, "%Y-%m-%d") + datetime.timedelta(days=1)).strftime("%Y-%m-%d")

    # Fetch historical SPY data for the desired date    

    # Fetch historical U.S. 10-Year Treasury Yield data
    treasury_data = yf.download("^TNX", start=desired_date, end=day_after_desired_date)
    treasury_yield = treasury_data.loc[desired_date, "Close"] / 100.0  # Convert to decimal

    # Calculate annual risk-free rate (0.01 = 1%)
    risk_free_rate = treasury_yield
    print("Stock price: ", stock_price)
    print("Risk Free Rate: ", risk_free_rate)
    print("Implied volatility: ", implied_volatility)
    print("Strike price:", strike_price)
    print("Time to expiration:", time_to_expiration)

    # Calculate put option price using the Black-Scholes formula
    put_price = BS_PUT(stock_price, strike_price, risk_free_rate, time_to_expiration, implied_volatility)

    return put_price


def calculate_profit(desired_date, desired_date_end, start_price, end_price, implied_volatility):
    # Parameters
    time_to_expiration = 5 / 365  # Time to expiration in years (6 months)
    number_contracts = 1
    strike_price_percentage = .01
    stock_price = start_price
    strike_price = stock_price - stock_price * strike_price_percentage

    # Calculate the put option price using the function
    put_price_start = calculate_put_price(start_price, implied_volatility, time_to_expiration, strike_price)

    print("\n\n\nTheoretical put option price: ", put_price_start)
    print("Number of Put contracts sold: ", number_contracts)
    print("Strike price: ", strike_price)
    print("Premium collected: ", number_contracts * 100 * put_price_start)
    print("Current stock price: ", stock_price)
    print("Break even stock price: ", strike_price - put_price_start)

    date_gap_days = (datetime.datetime.strptime(desired_date_end, "%Y-%m-%d") - datetime.datetime.strptime(desired_date,
                                                                                                           "%Y-%m-%d")).days

    print("elapsed time to end flag: ", date_gap_days)
    if date_gap_days < (time_to_expiration * 365):
        print("Buying back put")
        # buy back put option
        put_price_end = calculate_put_price(end_price, implied_volatility, time_to_expiration - (date_gap_days / 365),
                                            strike_price)
        profit_per_contract = (put_price_start - put_price_end)
        # calculate profit
        profit = profit_per_contract * number_contracts
    elif end_price < strike_price:
        print("Buying stock")
        # calculate profit
        profit = (end_price - strike_price) * 100 + (number_contracts * 100 * put_price_start)
    else:
        print("Contact expired")
        # calculate profit
        profit = number_contracts * 100 * put_price_start

    print("Profit: ", profit)
    print("\n\n____________________________________________\n\n")
    return profit


# Read the spreadsheet using pandas
df = pd.read_excel("algo.xlsx", sheet_name="Sheet1")

# Iterate through each row of the dataframe
for index, row in df.iterrows():
    # Get the desired date from the "Up flag date time" column and end date from "End flag date time"
    desired_date = row["Up flag date time"].strftime("%Y-%m-%d")
    desired_date_end = row["End flag date time"].strftime("%Y-%m-%d")
    volatility = get_volatility_for_dates(
        "https://www.alphaquery.com/stock/SPY/volatility-option-statistics/30-day/iv-put", desired_date,
        desired_date_end)
    print(volatility)
    start_price = row["Indicative price at Up flag"]
    end_price = row["Indicative price at end flag"]
    # Add the calculated profit to the "Profit" column
    df.at[index, 'Profit'] = calculate_profit(desired_date, desired_date_end, start_price, end_price, volatility)

# Save the modified dataframe back to algo.xlsx
df.to_excel("algo.xlsx", index=False)