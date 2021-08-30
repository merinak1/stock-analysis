# Stock Analysis 
 
## Overview of Project
The code developed to analyze year 2017 and 2018 is computing data correctly. However, the stock analysis code is reading data multiple times even though it was not necessary. Therefore this code possibily will not work or will be slower to compute if the sheet contains thousands of stocks. In order to overcome potential performance issue, the stock analysis code was refactored for better performance while ensuring the functionality remain uncompromised. 

## Results
The analyis shows that stocks in 2017 had better return rate than in 2018. Only TERP had negative return in 2017. In contrast, only ENHP and RUN had positive return in 2018. RUN is only stock that performed better in 2018 than 2017. 

There is clear difference between the code before and after it was refactored even though functionality and results remained unchanged. While the values on the Excel sheet is same before and after refactoring, the time taken to run the analysis reduced significantly. 

### Prior to refactoring code 
Prior to refactoring, it took 0.9909668 seconds to run 2017 analysis and took 1.030029 seconds for 2018 analysis. 

![myimage-alt-tag](/Resources/StockAnalysis_2017.png)
![myimage-alt-tag](/Resources/StockAnalysis_2018.png)

### After refactoring code

After refactoring, it took 0.1850586 seconds to run 2017 analysis and took 0.2259521 seconds for 2018 analysis.   

![myimage-alt-tag](/Resources/VBA_Challenge_2017.png)
![myimage-alt-tag](/Resources/VBA_Challenge_2018.png)

The time taken to compute data has reduced significantly after refactoring, it will clearly improve performance and efficiency. 

## Summary

Refactoring code has lots of advantages as it will improve performance and efficiency, improve readability of code and shorten the length of code in some cases. But at the same time, refactoring code has some disadvantages. It will not change the results generated but may still take lot of time to refactor. If the code is complicated, there is risk of inadvertently ruining perfectly running code. Some cases the time spent in refactoring code might not improved the efficiency and performance.  

In regards to refactored VBA script, the biggest advantage is that reduced time to complete running script. It have improved more than 50%. The disadvantage is that it did not add any new functionality to make it better than prior to refactoring. 

