# Financial Modelling Using MS Excel and VBA (JEM128)
The repo is for the project of Financial Modelling Using MS Excel and VBA from [IES](https://is.cuni.cz/studium/predmety/index.php?do=predmet&kod=JEM128) (Charles University) SS/2022.

# Triangular arbitrage
The goal for the project is to create first the data source, second the artitrage algorithm, and lastly an analysis of the arbitrage data. 
The triangular arbitrage algorithm takes advantage of the discrepancy between three foreign currencies when the currency's exchange rates do not match up.

## Getting the data
First I started with a list of the most used currencies that I found on the internet (See the pre-data.xlsm, tab raw) and modified it so I have just the tickers in column B.

Once all the data is preprocessed, we can use the algorithm GenerateCombinations, to create all the different combinations between the different currency tickers.

To get the latest price data for the currency pair, I used the built-in Excel price data with is powered by Bing. Select column A and Data-> Data Types-> Currencies. Now in column B just use =A1.Price to get the price.

![image](https://github.com/MyAscii/JEM128-IES/assets/68503801/53776153-7e30-4ba9-8fbf-74425f883abb)

## Running the script
Open arbitrage.xlsm, import your new data or use the old data, after you need to modify the path in visual basic code, and put the correct path where is the python file situated. After everything is done you can click to run the script. It will open a CMD command line and write some commands to run the Python script. If this seems scary to you, you can run the Python file separately. 

The output of the files is named Arbitrage_Opportunities.xlsx
