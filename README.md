# Stock-Analysis
Stock Analysis for Steve's parents

## Purpose and Background
The purpose of this project is to analyze the daily volume and return (in percent) of ‘green’ stocks for years 2017 and 2018 using a refactored VBA code. Refactoring the code was done to improve the efficiency of code previously written to speed up the amount of times it takes to analyze the daily volume and return of the stocks.  

The VBA code previously written to run analysis on green stocks works for the number of stocks being analyzed, but running the script on a large dataset like the entire stock market would take a long time. The VBA code was refactored to loop through all the data one time in effort to make it more efficient to make it run faster, use less memory, and make it easier for other users to read.

## Results
Using the refactored VBA code to analyze the daily volume and return of the twelve stocks for 2017 and 2018 showed an improvement in the amount of time it took to perform the analysis. For 2017, all stocks with the exception of one, TERP, returned a positive return for the year. TERP had a return of -7.2%. Four stocks (DQ, ENPH, FSLO, and SEDG) provided returned of over 100%. 
For 2018, all but two stocks had a negative return. ENPH and RUN had returns of 81.9% and 84.0%, respectively. Out of the ten stocks that had negative returns, six of them had double digit negative return. JKS had the largest negative return with a return of-60.5%.

With the original VBA code, it took .656 seconds (rounded) to run the analysis for 2017. Using the refactored code, it took 0.121 seconds (rounded) to run the analysis.

Original:

<img width="557" alt="VBA_Challange_2017_Original" src="https://user-images.githubusercontent.com/101379969/162601564-104a1a5a-0386-417b-9b2a-5551f7c05e3c.png">
Factored:

<img width="576" alt="VBA_Challange_2018_Original" src="https://user-images.githubusercontent.com/101379969/162601568-f1e7eec1-e6fe-4379-ac22-47616dd0c495.png">

With the original VBA code, it took .656 seconds (rounded) the analysis for 2018. Using the refactored code, it took 0.133 seconds (rounded) to run the analysis.

Original:

<img width="541" alt="Screen Shot 2022-04-09 at 9 24 34 PM" src="https://user-images.githubusercontent.com/101379969/162601574-af8d6afc-d451-4a9b-8bc7-e06a38ed9648.png">
Factored:

<img width="534" alt="VBA_Challange_2018_Refactored" src="https://user-images.githubusercontent.com/101379969/162601580-93c041a8-d998-4bbe-93bb-ec694f92747c.png">

In the original code, formatting was a separate macro. After running the analysis, an additional macro had to be run to format the cells to make it easier to identify which sticks returned a positive or a negative return for the analysis year. In the refactored code, it was integrated into the analysis so only one scipt had to be run.

![VBA_Original_Fomating_Macro](https://user-images.githubusercontent.com/101379969/162601951-339e5052-27b1-4219-833d-33ba1b21a789.png)

The formating code was moved to the end of the stock analysis portion of the refactored code before the code to stop the timer of the code.

![VBA_Refactored_Formating](https://user-images.githubusercontent.com/101379969/162601947-a564c314-b1d2-4910-afc5-43c6ccaecb7e.png)

For the analysis of the stocks, "tickerIndex" variable was created along with tickerVolumes, tickerStartingPrices, and tickerEndingPrices output arrays to simplify the code.

![VBA_Refactored_Analysis](https://user-images.githubusercontent.com/101379969/162602007-b753e90d-149d-4545-af8e-50fae8ecab54.png)


