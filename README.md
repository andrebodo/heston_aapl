# Heston Asset Pricing Model

An implementation of the Heston asset pricing model to determine an option's price. Using a dual-stocastic process to forecast prices while accounting for changes in volatility regimes, the price of they underlying is determined via a Monte-Carlo walk-forward method.

AAPL's stock data is pulled using Bloomberg Query Lanaugage in Excel (not shown due to Bloomberg connection requriement) and the maximum likelihood estimate (MLE) method is used to determine the Heston model parameters based on historical data.

Asset prices are generated using the Heston model and the option value at expriation is calculated and discounted to present to determine the price. A comparison to AAPL 4/29 120C ask price is included.

Code is not vectorized... because VBA. Carefuly when selecting the number of simulations since putting this number too high can crash excel.
