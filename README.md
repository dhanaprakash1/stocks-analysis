# stocks-analysis
Analyse Stocks using VBA

## Overview of the Project
### Purpose
Analyse how stocks performed for 12 tickers during years 2017 and 2018

## Results 
### Stock Performance
-  DQ Stock that we were intersted in did *NOT* perform well in 2018 even though it performed well in 2017.  
    - Investing on that stock based on 2017's results could have resulted in a loss
    - The Daily volume increased in 2018 compared to 2017
- ENPH and RUN performed well in both years and dail volume increased in 2018 compared to 2017

### Code Performance or Execution Time 
- The execution time for code written using class material was ~37 seconds
- The execution time for refactored code in challenge is ~0.13 seconds
- We iterated  through all the 3012 Rows 12 times, once for each ticker in classwork. The refactored code avoided that and the rows were iterated through only oncee for all 12 tickers
- Also, the sheet was activated for evey iteration in classwork, which could have contributed to additnal execution time
- [ExecutionTime_BadPerformance2018](Resources/VBA_Classwork_withBadPerformance.png)
- [ExecutionTime_RefactoredPerformance2018](Resources/VBA_Challenge_2018.png)
- [ExecutionTime_RefactoredPerformance2017](Resources/VBA_Challenge_2017.png)

##Summary 
### What are Advantages or Disadvantages of refactoring code
#### Advantages
- Comments could be added and unneccesary variables and debugging statements can be removed 
- Performance can be improved 
#### Disadvantages 
- The code has to be tested throroughly after any change. this is risky and time consuming 
- Its hard to recollect and change code that was written earlier, evebn if its just 2 weeks

### How to these pros and cons apply to refactoring the original VBA script?
#### cons
- the older working version which was not refactored had to be saved carefully as I doid not use git for version control while developing 
- Took a long time to recollect the code that was written 2 weeks ago
- Had to test the code once again thoroughly and it costed time
- refactoring is boring unless, it is something that is very interestign to do like improving performance like how we did
#### pros
- Refactoring helped in finding issues with already written code and fix it. These issues could have escaped, if refactoring was not done. 
- Performance improved 
- Helped wih learning as the same excercise was repeated in a different way



