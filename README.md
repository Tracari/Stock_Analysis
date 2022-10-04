# Stock_Analysis

## Overview
VBA is a programming language that interacts directly with Excel worksheets and cells, allowing us to write scripts to automate simple tasks, allowing even more analytical power to Excel. VBA is often used in the financial industry, where Steve just graduated with his Finance degree, and his parents are his first customers curious about investing. Though his clients are primarily interested in DAQO New Energy corporation (DQ), Steve wants to diversify and analyze other green energy stock options against DQ to offer the client variety. Using VBA, in this module, a macro will be created that can trigger pop-ups and inputs, read and change cell values, and format cells to consolidate and compare 12 stock options.

### Data Environment:
- VBA
- Microsoft Excel





## Results 

<img width="774" alt="Screen Shot 2022-10-04 at 8 33 06 AM" src="https://user-images.githubusercontent.com/105556091/193833117-84f8ce6b-0938-4667-b446-2a1cf6563e7d.png">


Stock performance: 
- DQ, the stock interest for this project, was the highest performing stock in 2017 (199.4%) but had the most vast performance pivot in 2018 (-62%). 
- RUN’s rate of return increased 78.55% making it the greatest increase when all the other stock were failing
- ENPH, one of the top performing stocks (129.5%) in 2017 fell in 2018 but was able to maintain a positive ROR performance, indicative of a well founded stock option.
- TERP’s rate of return was down (-7%) in 2018 and fell even lower in 2017 (-5%), denoting a negative return from YTD and low purchase interest

Code Performance:

The original code for AllStockAnalysis() ran in:
- 0.26 seconds: 2017
- 0.37 seconds: 2018

<img width="756" alt="Screen Shot 2022-10-04 at 8 34 17 AM" src="https://user-images.githubusercontent.com/105556091/193833472-8f1a9922-22e9-4854-af7f-f9cd7714e6ec.png">


The AllStockAnalysisRefactored() ran in:
- 0.08 seconds: 2017
- 0.07 seconds: 2018

<img width="729" alt="Screen Shot 2022-10-04 at 8 35 29 AM" src="https://user-images.githubusercontent.com/105556091/193833657-692b3fcd-6993-4fb5-8534-367fdb5a271b.png">


- The refactored code reduce processing time by creating empty arrays to hold data and nested IF:Then statements that looped through the data one time and collected all of the information for simultaneous calculations. verses the original code that ran the calculation for each ticker 1 by 1 
<img width="294" alt="Screen Shot 2022-10-04 at 1 36 57 AM" src="https://user-images.githubusercontent.com/105556091/193833825-f1084568-50da-4fa0-95f8-968bb90658da.png">
<img width="512" alt="Screen Shot 2022-10-04 at 1 34 11 AM" src="https://user-images.githubusercontent.com/105556091/193833910-e34b3d93-51a5-4eb4-9a7a-cdf0797c97ed.png">



## Summary 
In a summary, There are advantages to refactoring code:
- Improves functionality of code by removing redundancies and duplications
- Executes code faster
- Organizes code better

-And disadvantages to refactoring code:
- could introduce new bugs and errors into the code

For this exercise, Refactoring the original code made it’s organization more concise, Improved the processing times, and Improved the functionality for future testability.
