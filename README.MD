# Stocks Anlaysis with Excel And VBA 

## Overview of Project

This worksheet contains the information of 12 different stocks such as:
1.- Date
2.- Opening Price
3.- Higher price
4.- Lower Price
5.- Closing price
6.- Volume of the day

The objective of this analysis is to automate the calculation of the total volume obtained by each of the stocks in a given year (2017 or 2018).
The percentage of profit obtained by each stock is also calculated and the cell format is modified depending on the result obtained.

Finally, a button is added to run the Macro and the user can choose which year he wants to perform the analysis through an inputbox


### Purpose & Backgroung

One of the main objectives of this work is not only to calculate the parameters described in the previous section, but also to do it in an efficient way, optimizing the code so that it is executed using the least amount of computational resources.

 ## Nested for loop

The first approach to solving this problem was to use a nested for loop within another loop.

The first for loop assigns the value of a stock. example:
Stcok (0) = "AY"
 
And then the nested for loop loops through each box within the dataset. therefore, to solve the problem, the algorithm needs to go through the entire worksheet 12 times, which consumes time and resources.


## Single for Loop

In order to optimize the algorithm, only a for loop is used to go through each of the rows that contain the data of interest.
Basically the algorithm obtains the beginning and the end of each stock and performs the necessary arithmetic operations for the data we are looking for.

A second for loop is used only to give a frame to the cells where the information will be displayed.

If the return is positive, the cell will be filled with green, while if it is negative, it will be filled with red.

## Results
The analysis of results is divided into 2 parts, The analysis of the stocks to determine if it is a good investment option.
And algorithm analysis to determine if there is a substantial improvement in execution times or not.
### Stocks Analysis 
2018 seems not to have been a good year for the markets in the green energy industry, since only 16% of the total representative sample reported profit in 2018.

"DQ" reported the worst fall within the stocks analyzed, dropping the price of its stock by more than 62%.

Although "RUN" reported profit both years, the best option to invest is "ENPH" which has shown the highest profit in both years.

### Run time Analysis
The data of the time it took for the computer to execute the algorithm for calculating the stocks is presented using a nested for loop VS a single for loop.

For the nested for loop the running time was = 2.195313 for the year 2017 and 2.1523 for 2018. 

For a single for loop the runnig time was = 0.2734375 for 2017 and   0.3281 for 2018

This means tha a single for algorithm solves the problem almost 811% faster than a nested for loop

## Summary
It is difficult to conclude that one algorithm is better than another based solely on its running time.

The single For loop algorithm works only because we know from the beginning that the data is perfectly organized by day.

For this specific case, the use of a single for loop was the most optimal way to solve the problem, but it does not mean that it will always be the case. It depends on the organization of the data that we want to analyze.