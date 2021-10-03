# Green Energy Stock Analysis 
## Overview
#### Our good friend Steve just recently graduated from college with his Finance degree. His parents, are going to be his first clients. Steve’s parents are passionate about green energy yet they are not really educated on how to wisely invest their money in green energy stocks. Currently, they have chosen to invest all their money into DAQO New Energy Corp. Steve is not confident in this decision and has asked us to do some analysis on DAQO and other comparable green energy stocks. He has pulled some data for the stocks he is looking to analyze and has provided that data to us in an excel file. 
#### In the following, we will run analysis for DAQO (Ticker: DQ) for the year 2018 per Steve’s request as well as create a more comprehensive review of all the stocks that Steve has chosen that he can run for each year he has data. Both analyses will look at the total daily transactional volumes as well as the rate of return. 
---
## Results: 
#### Using the dataset that was provided to us by Steve in an excel file, we will create macros using VBA to build tools that Steve can easily use to run analysis on stocks to help guide his parents to better financial decisions. Previously, we built macros for Steve that would perform his needed analysis. After ensuring that the information was accurate and that Steve could help use it to better invest his parents hard earned money we tuned our focus to improving the overall performance of the macros we created. In the following, we will first, compare the overall financial results year over year and then go into detail on how we improved the performance of the macros. 
---
### Financial Results Year Over Year
#### Below are two images. One showing the return of all the stocks listed in 2017 and another showing the return for the same stocks in 2018. Using color formatting of green for a positive return and red for a negative return, it is easy to identify that the market overall had a better return in the year 2017 than that of 2018. It was also easy to identify that only two stocks had positive returns in the year 2018, those stocks carried the ticker “ENPH” and “RUN”. Based on the results in 2017 and 2018 the stock with the ticker “ENPH” appears to have the best overall return. 
---
![VBA_Challenge_2017](https://user-images.githubusercontent.com/90698381/135774388-ecfb1aab-2390-4010-9f7f-1a6982a59ce0.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/90698381/135774396-0ca30fa7-54c4-437c-a534-a7b85be92aa9.png)
---
### Improving Macro Performance
#### Our initial macros ran at a reasonable speed but Steve was somewhat concerned that if he were to evaluate a larger dataset that the performance would be subpar.  He also asked us to try and build the formatting into the initial analysis button we created to streamline his process. We created a new macro, starting with the code we used for Steve’s All Stock Analysis. We removed the formatting button and added it to the new macro. We then refactored the code we created for our initial macro to significantly improve its performance. One of the key improvements to speed of the performance was to eliminate the nested for loops that were originally contained in our macro. The image below shows an example of our initial code. 
---
![DQCode](https://user-images.githubusercontent.com/90698381/135774404-f7346f8c-810e-449f-ae4a-e623e29bb5c4.png)
---
#### While the above code is simple it does slow performance. By removing the nested loop and creating individual for loops the performance was increased. The two below images show the individual for loops that were created. 
---
![ASAR1](https://user-images.githubusercontent.com/90698381/135774415-9040c67d-ab3f-46a0-ad0f-838f6432974d.png)
![ASAR2](https://user-images.githubusercontent.com/90698381/135774424-9b6b0e02-75e6-40e5-a045-4c56e6f4d1fe.png)
---
#### The following screens shots show the completion time for the initial macros for each year.
---
![2017_Original](https://user-images.githubusercontent.com/90698381/135774433-edae0a41-e64f-481d-a6e1-81cf607eb8c6.png)
![2018_Original](https://user-images.githubusercontent.com/90698381/135774436-038e7ad5-40ab-4150-89b1-98f3755c1f0c.png)
---
#### The next two images show the performance of our macro after the code was refactored. Please note that it runs in a fraction of the time while including the formatting that once required a separate button. 
---
![VBA_Challenge_2017](https://user-images.githubusercontent.com/90698381/135774447-4abe5d93-5075-4b09-b6d0-10abe937c3ef.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/90698381/135774449-f1af9241-01a0-47d9-b0ea-6089072fe70f.png)
---
## Summary: 
### While our code ran at a reasonable time there are clearly advantages to having refactored it. The first advantage is that it runs much more efficiently and in a fraction of the time of our original macro. Also having refactored it and including the formatting into one routine we have eliminated a few extra clicks for Steve. 
### The downside to refactoring the original VBA script is that there are a few more lines of code to read though and if any future adjustments are needed it may be more complicated to implement them. 
