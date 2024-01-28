
"""
The API I used here is ExchangeRate-API (https://www.exchangerate-api.com/).
"""

import requests
import pandas as pd
from datetime import datetime

import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)


class ExchangeRateAPI:
    """
    This class is used to get the exchange rate data from the API
    """

    def __init__(self, api_key):
        self.api_key = api_key
        self.base_currency = 'USD'
        self.target_currencies = ["AUD", "CAD", "CHF", "CNY", "EUR", "GBP", "HKD", "JPY", "NZD", "USD"]
        self.api_url = f'https://v6.exchangerate-api.com/v6/{self.api_key}/latest/{self.base_currency}'

    def fetch_exchange_rates(self):
        """
        This function gets the exchange rate data from the API
        """
        try:
            response = requests.get(self.api_url)
            response.raise_for_status()

            data = response.json() # convert the response to JSON
            if data['result'] == 'success':  # if the response is successful
                conversion_rates = data['conversion_rates']  # get the conversion rates in a dictionary
                return conversion_rates
        except requests.exceptions.RequestException as e:  # check for any request exceptions
            print("Error fetching exchange rates: {}".format(e))
        except ValueError as e:  # check for any value errors
            print("Error parsing the response: {}".format(e))
        except Exception as e:  # check for any other exceptions
            print("Error: {}".format(e))

    def save_to_excel(self, exchange_rates, output_file):
        """
        This function saves the exchange rates to an Excel file
        """

        current_date = datetime.now().strftime("%Y-%m-%d") # get the current date for the excel file

        data = {
            "Rate Type": ["Spot rate"] * len(self.target_currencies),
            "Date": [current_date] * len(self.target_currencies),
            "Currency_From": [self.base_currency] * len(self.target_currencies),
            "Currency_From_Value": [1] * len(self.target_currencies),
            "Currency_To": self.target_currencies,
            "Currency_To_Value": [exchange_rates[currency] for currency in self.target_currencies]
        }

        df = pd.DataFrame(data)  # create a dataframe from the data
        df.to_excel(output_file, index=False) # save the dataframe to an excel file
        print("Saved the exchange rates to the file: {}".format(output_file))

    def generate_exchange_rate_table(self, output_file):
        """
        This function generates and saves the exchange rate table to an Excel file for all the target currencies
        """
        exchange_rates = self.fetch_exchange_rates() # get the exchange rates from the API
        if exchange_rates: # if the exchange rates are not empty
            self.save_to_excel(exchange_rates, output_file) # save the exchange rates to an excel file
        else:
            print("No exchange rates found")