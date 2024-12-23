                                                   Analysis Report
                             Prepared for Cryptocurrency Data Fetching and Analysis Task

Objective

The purpose of this task is to fetch live cryptocurrency data for the top 50 cryptocurrencies, analyze it, and present the data in a live-updating Excel sheet.

Data Fetched:
The script retrieves live data for the top 50 cryptocurrencies based on market capitalization, including the following key fields:
1. Cryptocurrency Name
2. Symbol
3. Current Price (USD)
4. Market Capitalization
5. 24-hour Trading Volume
6. 24-hour Percentage Price Change

Key Insights
1. Top 5 Cryptocurrencies by Market Capitalization:
    Based on the latest data, the top five cryptocurrencies are determined dynamically using
    market capitalization as the key metric.               

2. Average Price of the Top 50 Cryptocurrencies:
    The script calculates the average price of the top 50 cryptocurrencies, giving an overview of 
    the market's pricing trends.

3. Highest and Lowest 24-hour Percentage Price Change:
   The cryptocurrency with the highest percentage price gain in the last 24 hours is highlighted, 
    providing insights into market momentum.
   Similarly, the cryptocurrency with the largest percentage price drop is identified.

 Methodology
1. Data Source:
   The script fetches live data using the CoinGecko API.

2. Data Analysis Techniques:
    Pandas library is used to process and filter the dataset.
   - Insights are extracted through sorting, aggregation, and selection of relevant metrics.

3. Excel Integration:  
The processed data is stored in an Excel sheet using the OpenPyXL library. The sheet               updates every 5 minutes to reflect the latest data.

Below is a sample result from a recent execution:


Conclusion:
The live cryptocurrency data highlights dynamic price fluctuations and market trends.
The analysis provides actionable insights, such as identifying top-performing assets and market-wide pricing behavior.
The continuous Excel update ensures that the data remains relevant for timely decision-making.
