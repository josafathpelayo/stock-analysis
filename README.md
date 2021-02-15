##Josafath Pelayo

# stock-analysis

## Overview
The purpose of this analysis is to refractor a VBA code used to loop through a dozen stocks in module 2 so it can handle running  an extensive list of stocks.Origannly, the code was set up to run an analysis on 12 stocks, but by refractoring, the speed and code overall can be enhanced. Then it was determine whether refactoring the code successfully made the VBA script run faster. When finally ruuning our new VBA script on the same data set for Module 2, pop-ups with script run time for both new and old were used to compare their effeincy. Essentially, the speed of the codes was the varaible used to determine if refactoring was successful or not. 

## Results
When running my refractored code [Refractored](https://github.com/josafathpelayo/stock-analysis/blob/main/Screenshot%20(76).png), [Refractored#2](https://github.com/josafathpelayo/stock-analysis/blob/main/Screenshot%20(77).png), I was able to run a time of 0.1914063 for 2017, where as 2018 ran at 0.1835938 seconds. These times where were seen in the message box when the script was ran,[2017msgBox](https://github.com/josafathpelayo/stock-analysis/blob/main/VBA_Challenge_2017.png), [2018msgBox](https://github.com/josafathpelayo/stock-analysis/blob/main/VBA_Challenge_2018.png). When the original code was ran [Original_Code](https://github.com/josafathpelayo/stock-analysis/blob/main/Screenshot%20(78).png) [Orignal_Code#2](https://github.com/josafathpelayo/stock-analysis/blob/main/Screenshot%20(79).png), both 2017, and 2018 had a time of 1.441406 seconds [2017_Orignal_time](https://github.com/josafathpelayo/stock-analysis/blob/main/2017%20old%20code%20run%20time.png) [2018_Orignal_time](https://github.com/josafathpelayo/stock-analysis/blob/main/2018%20old%20run%20time.png). The resutlts of both codes gave the ouput of [2017_Chart](https://github.com/josafathpelayo/stock-analysis/blob/main/2017%20Stocks%20chart.png) and [2018_Chart](https://github.com/josafathpelayo/stock-analysis/blob/main/2018%20Stocks%20chart.png). The evidence of cutting the run time by refractoring the VBA script proves that refractoring does make a big difference when writing code. 


## Summary
### Advatages and Disadvantages of Refractoring
When refracting a code, the upside can be great if theres a high chance of enhacing the code. For example, condensing codes if the script looks long, has redundancies, or has too many bugs. If it looks wrong and can be enhabced, why not. Most of the time, just having another peer review one's work is refactoring. 

The main downside of refactoring is the proccess of refactoring itself! In a perfect world, refactoring codes to inhance the speed, or simply restructuring code without changing its behavior, creates a better code overall. The time consumption in doing so, can be costly in the real word when there's deadlines near. For example, if it takes a longer time to refactor a code rather than starting from scratch, do not do it. Another example is if the code itslef is already stable, or if you don't have time to test the new code. It all comes down to if it's worth the time or not. 

Overall, when it came down to refacting the original code, it had it advantage as it did improve our run time. For our old code run time, both 2018 and 2017 ran at 1.441406 seconds. For our new run time, 2017 ran at 0.1914063 where as 2018 ran at 0.1835938 seconds. The length of the code was also an advantage as a longer and more complex code would accumilate an extensive time. Although its good to mention that the processing speed of the computer being used can be an external factor that effects the run time. 
###
