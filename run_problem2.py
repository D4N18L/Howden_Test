
from load_config import load_config
from Problem2 import ExchangeRateAPI


def main():

    config = load_config()

    # Get the API key from the config file
    api_key = config['api_key']

    # Pass the API key through to the ExchangeRateAPI class
    exchange_rate_api = ExchangeRateAPI(api_key)

    # Generate the exchange rate table and save it to an Excel file
    exchange_rate_api.generate_exchange_rate_table("problem2_exchange_rate_table.xlsx")

    print("Finished generating the exchange rate table.")

if __name__ == "__main__":
    main()
