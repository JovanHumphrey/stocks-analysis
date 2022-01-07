#**Stock Analysis Using Excel and VBA**

##**Overview**

My client asked that I create a script to help an end-user review information for a list of stocks in 2017 and 2018. Having created a functional script for the provided dataset I refactored the code to improve run time.

##**Purpose**

At the click of a button the end-user should be able to review the daily volume and return percentages of various stocks on a given year. With that information the user should be able to choose better-performing stocks over less optimal stocks.

My initial script [AllStocksAnalysis()] worked well on 3013 rows of data for twelve stocks but had an unreasonably long runtime.

! [2017 Runtime](stocks-analysis/2017 runtime.PNG at main · JovanHumphrey/stocks-analysis (github.com))

! [2018 Runtime](stocks-analysis/2018 runtime.PNG at main · JovanHumphrey/stocks-analysis (github.com))

This would be untenable if the script was repurposed for a longer list of stocks. If left unchanged, this script could slow the user’s computer down considerably. I needed to refactor the script to decrease the runtime.

###**Original Script All Stocks Analysis

In the original script the code was designed to run through 3013 rows of data to collect information about twelve different stocks. With the use of nested for-loops and if/then conditions, the code did two primary things:

-1. It identified the starting and ending price for each stock and calculated the rate of return based on the difference.

- 2. It added up the total daily volumes for each stock. See the results below.

![2017 Results](stocks-analysis/2017.PNG at main · JovanHumphrey/stocks-analysis (github.com))

![2018 Results](stocks-analysis/2018.PNG at main · JovanHumphrey/stocks-analysis (github.com))

###**Refactored Script All Stocks Analysis Refactored

In order to run the information, the original script utilized nested loops. This led to a slower runtime.  To correct this, I modified the if/then conditions to recall information in the array I had previously declared. While the original script had to run the if/then conditions of each stock before looping, the refactored script told the code to run the if/then condition while recalling the values in the array. The result was a visibly faster experience.

You can view my original and refactored scripts [here](stocks-analysis/VBA_Challenge.xlsm at main · JovanHumphrey/stocks-analysis (github.com))

##**Result**

Refactoring allowed my script to run hundreds of thousands of times faster than the original.

![Refactored 2017 Runtime](stocks-analysis/refactored 2017 runtime.PNG at main · JovanHumphrey/stocks-analysis (github.com))

![Refactored 2018 Runtime](stocks-analysis/refactored 2018 runtime.PNG at main · JovanHumphrey/stocks-analysis (github.com))

##**Summary**

###### **Advantages and Disadvantages of Refactoring**

In general, refactoring code allows the loop to perform fewer steps with each row of data—allowing the computer to do the same work more efficiently.  It also has the benefit of being easier to read and easier to expand upon in the future. Because it contains fewer lines of code refactoring can also make it easier to spot errors.

However, there are a few drawbacks to refactoring: it can be time consuming. The amount of time it takes to refactor a code may not yield significant savings in runtime. Additionally, you must adjust the logic of the code to eliminate unwanted loops. This can be confusing and difficult to keep track of.

######**Pros and Cons of Refactoring This Project**

The pros of refactoring the Stock Analysis script far outweigh any cons. My original script took several seconds and caused my Excel window to flicker while it ran. The refactored version runs quickly without any glitches. The only drawback I found was as a novice programmer I didn’t know how to best utilize my array to simplify the code. It took a lot of trial and error and Googling to find a solution that worked. Had the solution not yielded such excellent results, it wouldn’t have been worth the hours I spent working on it.
