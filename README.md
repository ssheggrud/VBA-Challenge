#The VBA of Wall Street

##Project Overview
The client wishes to be able to review the entire stock market over the last few years. Due to the large dataset, the original macros would run very slowly. My task was to refactor the code so that it would run faster. Additionally, I needed to make sure that there were no changes to the results which could indicate an error in the code.

###Results:
After refactoring the code, I was able to reduce time processing time to approximate one-third of the original.

![Process time for the 2017 Stock Data Analysis](https://github.com/ssheggrud/VBA-Challenge/blob/main/resources/VBA_Challenge_2017.png)

![Process time for the 2018 Stock Data Analysis](https://github.com/ssheggrud/VBA-Challenge/blob/main/resources/VBA_Challenge_2018.png)

##Summary
Because I wanted to be able to compare the original code, both runtime and results, to the refactored code, I chose to create a new worksheet for the refactoring. I named it "Challenge".

The biggest disadvantage I encountered when refactoring the code was figuring out why an overflow error was occurring. With help of the TA, I was able to fix the issue by removing an unnecessary step.

In comparing the advantages/disadvantages of the original vs the refactored data, I really don't see an advantage to using the original code. It is slower and will become noticeably sluggish as more data is added in. With a very small dataset, the improved processing time of the refactored code is nearly imperceptible.
