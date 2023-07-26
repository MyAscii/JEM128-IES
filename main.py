import importlib

def install_package(package):
    try:
        importlib.import_module(package)
        print(f"{package} is already installed.")
    except ImportError:
        import sys
        import subprocess

        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"{package} has been installed.")

# Install pandas
install_package("pandas")

# Install openpyxl
install_package("openpyxl")

# Install itertools
install_package("itertools")

import pandas as pd

import os

actualpath = os.path.dirname(os.path.abspath(__file__))

file_name = 'Arbitrage.xlsm'
file_path = os.path.join(actualpath, file_name)
df = pd.read_excel(file_path)

# Perform triangular arbitrage
def triangular_arbitrage():
    opportunities = []
    
    # Define and populate the conversion_rates dictionary
    conversion_rates = {}
    
    # Extract currency pair and price information from the DataFrame
    for index, row in df.iterrows():
        ticker_pair = row['Ticker Pair']
        price = row['Price']
        base_currency, quote_currency = ticker_pair.split('/')
        conversion_rates[(base_currency, quote_currency)] = price

    
    # Iterate over all possible triangular arbitrage opportunities
    for (base_currency, quote_currency) in conversion_rates.keys():
        for intermediate_currency in conversion_rates.keys():
            if intermediate_currency != (base_currency, quote_currency):
                if (quote_currency, intermediate_currency[1]) in conversion_rates and (intermediate_currency[1], base_currency) in conversion_rates:
                    cross_rate = (conversion_rates[(quote_currency, intermediate_currency[1])] / 
                                  conversion_rates[(intermediate_currency[1], base_currency)])
                    if cross_rate > conversion_rates[(base_currency, quote_currency)]:
                        opportunity = {
                            'Base Currency': base_currency,
                            'Quote Currency 1': quote_currency,
                            'Quote Currency 2': intermediate_currency[1],
                            'Rate 1': conversion_rates[(base_currency, quote_currency)],
                            'Rate 2': conversion_rates[(quote_currency, intermediate_currency[1])],
                            'Rate 3': conversion_rates[(intermediate_currency[1], base_currency)],
                            'Cross Rate': cross_rate
                        }
                        opportunities.append(opportunity)
                        print(opportunity)
    
    # Create a DataFrame from the opportunities list
    opportunities_df = pd.DataFrame(opportunities)
    opportunities_df = opportunities_df.drop_duplicates(['Base Currency', 'Quote Currency 1', 'Quote Currency 2'])

    # Convert the cross rate to the base currency
    for index, row in opportunities_df.iterrows():
        opportunities_df.at[index, 'Cross Rate'] = 1 / row['Cross Rate']
        
        # Calculate the profit percentage
        profit_percentage = ((100000000 * row['Rate 1'] * row['Rate 2'] * row['Rate 3']) / 100000000) 
        opportunities_df.at[index, 'Profit Percentage'] = profit_percentage
    
    opportunities_df = opportunities_df.sort_values('Profit Percentage', ascending=False)

    
    return opportunities_df

# Run the triangular_arbitrage function and print the opportunities DataFrame
arbitrage_opportunities_df = triangular_arbitrage()
print(arbitrage_opportunities_df)


file_name = 'Arbitrage_Opportunities.xlsx'
file_path = os.path.join(actualpath, file_name)

arbitrage_opportunities_df.to_excel(file_path, index=False)

