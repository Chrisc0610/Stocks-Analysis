# Kickstarting with Excel

## Overview of Project
To help Steve's parents better analyze the entire stock martket data over a given number of years. While also optimizing the current code that was created for steve.
In order to get the most efficient time frame and data set I had to break down the set of tasks in the following order:
	- Refractor the code.
	- make sure to cut time in processing of code. 
	- make sure it loops thorugh all rows for the selected year and ticker.
	- Make sure the conditionals work properly so that the data can be read fluently.

## Results
Based of the results it looks like 2017 was a much better year than 2018 when it comes to the 12 selected stocks. 
Luckily our new refractor code can help Steves parents choose a new set of stocks to invest on.
	- [2017 Stocks Outcomes](https://github.com/Chrisc0610/Stocks-Analysis/blob/main/Resources/all_stocks2017.PNG?raw=true)
	- [2018 Stocks Outcomes](https://github.com/Chrisc0610/Stocks-Analysis/blob/main/Resources/all_stocks2018.PNG?raw=true)

Once the refractor code executed the first test I ran was to verify that the code runs with no errors and that the processing times upgraded.

- In order to achieve this the orginal code had to be cleaned up so that it can run in a more organized fashion. 
- The following snips of code below shows how i acheived to cut down the run time.
	* I re organized the nested loops as well as added 4 new variables "tickerIndex, TickerVolumes, TickerStartingPrices, TickerEndingPrices" 
	* Then set the to 0 to be looped over.
	* the image below show a sample of this.
- [SnipofCode1]https://github.com/Chrisc0610/Stocks-Analysis/blob/main/Resources/SnipofCode.PNG?raw=true)
	
	
- Then I created the nested loops that will run through all rows of the selected year.
	* These conditionals will increase the TickerVolumes per row, check if it is in the first row or last rown and add the price to its perspective variable 
	* The image below show a sample of this
	- [SnipofCode2]https://github.com/Chrisc0610/Stocks-Analysis/blob/main/Resources/SnipofCode2.PNG?raw=true)
				
- We can tell by the images below that the code succesfully ran and in a more optimized time.
	
	
## Summary
The advantages on refractoring the code is that it will help the processing time of the code especially if there is more data added for more years in the future. 
We can also now use this code to run on another data set of stocks now that the it has been refractored. A disadvantage of refractoring a code is that there is a 
possibility that while refractoring the code, the old code will become unstable and not run properly. this is why it is best to try to run on two copies of your data 
to properly see the changes and updates while refractoring the script.

The advatages of the original VBA script was that while it took longer to run, the code was easier to read especially for begginers. the disadvantage of the original code was that
since it was written inh a more simple way, in the process of refractoring the code, alot of the code will have to be re-written which can be a challenge for beginners.

The advantage of the refractored script is that now it can be edited and run for a different dataset, as well as adding more years for a greater comparison, by just edited some of the 
conditionals. the disadvantage of the refractored VBA script is that due to all of the nested loops, there can be a room for errors if the variables are not properly called. 
