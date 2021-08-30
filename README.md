# Stock Analysis 
 
## Overview of Project
The code developed to analyze year 2017 and 2018 is computing data correctly. However, the stock analysis code is reading data multiple times even though it was not necessary. Therefore this code will not possibily work or will be slower to compute if the sheet contains thousands of stocks. In order to overcome potential performance issue, the stock analysis code was refactored for better performance while ensuring the functionality remain uncompromised. 

## Results
There is clear difference between the code before and after it was refactored even though functionality remained unchanged. While the values on the Excel sheet is same before and after refactoring, the time taken to run the analysis reduced significantly.

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
Refactoring code has lots of advantages as it will improve performance, efficiency, improve the readability of the code and shorten code in some cases. But at the same time, refactoring code do not change the results generated but may still take lot of time to refactor. 
If the individual do not understand the code, there is risk of inadvertently ruining perfectly running code. 
