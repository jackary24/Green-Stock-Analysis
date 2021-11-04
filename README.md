# **VBA Green Stock Analysis**

---

## Overview of the Project
  In this module we were asked by a friend, Steve, to analyze a number of renewable energy stocks that his parents are looking to invest in. We were given the daily data for 12 different tickers, divided into the years 2017 and 2018, and asked to help them make a decision based on this information for the best return on their money. Originally they were interested in the stock with the “DQ” ticker, but after running our initial analysis, we were tasked with looking at all the tickers, as “DQ” dropped over 63% in the year 2018. Once we provided a suitable VBA macro to analyze all the tickers in the years 2017 and 2018, we were asked to refactor our codes in order to make them more efficient, and return the results in less time. 

---

## Results of Analysis

![2018 results](https://github.com/jackary24/Green-Stock-Analysis/blob/main/Visuals/2017_results.png) ![2017](https://github.com/jackary24/Green-Stock-Analysis/blob/main/Visuals/2018_results.png)

  In the year of 2017 we saw a generally positive return in the 12 stocks provided to us, with only one (TERP) actually having a negative return in that year. The highest rates of returns for the included stocks were seen in SEDG (184.5%) and DQ (199.4%). Had Steve’s parents invested early in the year of 2017 into their preferred DQ stock, they would have had a pretty good year. Unfortunately the following year of 2018 was not a good year for the green stocks, as all but two (ENPH & RUN) saw a negative rate of return. Both ENPH and Run would have been better investments as they both saw a rise in the return throughout their two years, as well as a large increase in their daily volumes.

---

## Differences in Code
  The largest difference between the original code and the refactored version is the speed in which they run. While the original code took about 0.7 seconds to completely iterate through the data set, the restructured VBA completed the task in a little over 0.1 seconds. You can see an example of this for the 2018 analysis listed below.

![2018 results](https://github.com/jackary24/Green-Stock-Analysis/blob/main/Visuals/2017_Original_vs_Refactored.png) ![2017](https://github.com/jackary24/Green-Stock-Analysis/blob/main/Visuals/2018_Original_vs_Refactored.png)

  The largest contributing factor to this decrease is found in the removal of the nested for loops, instead opting to separate the two, and using an array to store the output values. 

![0g code](https://github.com/jackary24/Green-Stock-Analysis/blob/main/Visuals/Original_code.png)

  As you can see above, the original iterated through our tickers in the (i) loop, and used the nested for loop (j) to determine the values from ticker ID, total volume, and the return. Once it calculated those three, it recorded them on the “All Stocks Analysis” worksheet before repeating the outerloop for the next (i) value.

![newcode](https://github.com/jackary24/Green-Stock-Analysis/blob/main/Visuals/Reformatted_code.png)

  Our new code (seen above) separates those for loops in order to have less iterations over the data set, insteading setting up the variables as arrays holding twelve values. We create an index, equal to 0 and increasing by 1 after each iteration, in order to loop through the arrays, allowing us to move the final output to a separate loop.

---

## Advantages & Disadvantages of Refactoring

### In General
  One of the major benefits to refactoring code is to improve the overall design of the code and decrease the amount of processing power necessary in order to run it. We saw this as a clear example in this module with the decrease in time the VBA took to execute. Additionally, refactoring can also increase the readability of your code by structuring it in a more logical progression.

  A downside to the refactoring process is the introduction of potential bugs. Restructuring the written code can oftentimes turn working code into non working code with a host of new error messages that require troubleshooting in order to address. This can seriously hamstring you if you are working on a tight deadline or budget.

### On the VBA
  The largest benefit of refactoring the module code during the challenge was making it run in less time, requiring less processing power. While the time difference between the two was not noticeable (other than in the pop up messages), if we were working with larger datasets, say the entirety of the US stock market, this change would be readily apparent. Additionally, refactoring helped me better understand how the code worked in general, and made me more comfortable working with VBAs.

  The Original VBA did. have one advantage over the refactored code however, although that was not readily apparent using the current dataset. As stated above, we removed a nested for loop, and used arrays in order to store data, instead of assigning them as we went. We used a total of 3 arrays, each holding 12 values each, for a total of 36 values. These are stored in the computer’s working memory, and we made the trade off between a lower speed and higher memory utilization. If we were to expand this code, say to look at the whole stock market, we would be storing a much larger number of values in the working memory, which could lead to errors. While writing code, one should be aware of this trade off, and find a balance between the two.
