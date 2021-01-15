# Stock Analysis
Module 2 VBA
During this analysis there were 12 green stocks analyzed for years 2017 & 2018. The analysis showed a year over year for trading and the ability to change the year by using macros. Steve, recently graduated with a degree in Financ and his first client were his parents.  They wanted him to look into DAQO stocks and analyze all the data so his parents can see a YoY on different trends. All findings for the project can be found here: [Green Stocks](https://github.com/EJones621/stock-analysis/blob/main/VBA_Challenge.xlsm)
## 2017 vs 2018
Here is an image based on how 2017 and 2018 did:


![2017 & 2018 Data](https://github.com/EJones621/stock-analysis/blob/main/Resources/2017%262018_Data.png)


As we can see, 2017 did much better than 2018 did.
ENPH did by far the best.
DQ had the biggest flip from highest to lowest.

##### What do the column headers mean?
Ticker: This indicates an abbreviation of the company name.

Total Daily Volume: The sum of the number of shares for trading days on the market.

Return: This indicates the Year over Year (YoY) which measures how the stock has performed over the course of the year. This can be found by: (last price / first price) - 1.


## Before Running New VBA Script
Let's take a look at the original VBA Script and see how quickly the data was able to run.
*This is done by clicking on the button "Run Analysis for All" on the "All Stocks Analysis" Worksheet*

![2017 Speed](https://github.com/EJones621/stock-analysis/blob/main/Resources/2017%20Speed%20Before%20Challenge.png)



![2018 Speed](https://github.com/EJones621/stock-analysis/blob/main/Resources/2018%20Speed%20Before%20Challenge.png)





## Speed of 2017 vs Speed of 2018 with NEW VBA Script
After updating the VBA Script to run quicker and more efficiently:

![2017 Speed](https://github.com/EJones621/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)



![2018 Speed](https://github.com/EJones621/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)




## Summary Statement
The advantages and disadvantages of the refactored may not be so clear with the difference between the old and new script being about one second. Both scripts do run quickly and efficiently, but the refactored one obviously runs just a little bit more quickly and efficiently. The old VBA Script may not run as smoothly if the dataset were bigger, while the refactored one wouldn't have issues.  The old VBA Script is a lot easier to create than the refactored, so if the dataset is small, I don't think it's necessary to run such complex coding, because it can get crazy.
