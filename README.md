# Stock_Analysis
Steve's parents stock


Overview

The purpose of this project was to learn how to work with mass amounts of data and create Macros to give us multiple steps of analyzing this data with the click of a button. We gave Steve's parents a look at not only the one stock they were particularly interested in but also compared it to 11 other stocks with data from two seperate years of activity for each stock. 

Results


First I had started with the original set of data and analyzed just the DQ stock that Steve's parents were interested in. I created a basic breakdown of the stock to show the daily volume and year return for 2018 of the DQ stock which showed that the DQ stock did not have a positive return for that year. 

<img width="222" alt="Screen Shot 2021-03-17 at 7 55 34 AM" src="https://user-images.githubusercontent.com/78769464/111470803-3ab5c580-86f6-11eb-9c0a-94fcb1e0d29e.png">


To do this I created a header row to make the information easily readable for Steve's parents. I had to create a command that pulled the amount of rows containing the data I needed to analyze. Using that selected data I used VBA to loop through each row pertaining to the DQ stock and calculated total volume for the DQ stock in 2018. The starting and ending prices that allowed me to seperate where the stock started in the beginning of the year and where it ended determined the total return for 2018.


<img width="665" alt="1- DQ Analysis" src="https://user-images.githubusercontent.com/78769464/111470894-50c38600-86f6-11eb-9385-ca900dd231bd.png">

After providing Steve's parents with an outcome for the 2018 DQ stock, I decided it was best to evaluate the other 11 stocks to allow his parents to have other options aside from the DQ stock. 

<img width="894" alt="Screen Shot 2021-03-17 at 8 41 24 AM" src="https://user-images.githubusercontent.com/78769464/111476996-ac910d80-86fc-11eb-8522-e5243104eab6.png">


![All Stocks Collage](https://user-images.githubusercontent.com/78769464/111475828-7bfca400-86fb-11eb-945d-989e1feb3183.jpg)


The process of analyzing the DQ stock versus the other 11 stocks was very similar with the minor differences of accounting for all 12 stocks. To do this I had to establish a process in which VBA looped through every row of each stock and calculated each volume, start price, and ending price to provide the information we provided for the DQ stock individually. Given the fact that each year has thousands of rows of data combined for all 12 stocks, I added a timer to see how long it would take to run the entire code. Then to make it easier to manipulate the data, I edited my analysis to be flexible for Steve's parents and to give them an option on how to select the year they wanted to run for all the stocks. Doing this I created a message box that initially pops up to have the user input the year they want and the sheet will populate the calculated data for that year.  To go even further to allow easy manipulation for his parents, I added user friendly buttons on the worksheet that allow Steve's parents to clear the worksheet and re-analyze which set of data they needed.

<img width="204" alt="Clear Worksheet Buttons" src="https://user-images.githubusercontent.com/78769464/111478914-7fddf580-86fe-11eb-9463-2a81f8d4e391.png">


<img width="618" alt="Clear Worksheet" src="https://user-images.githubusercontent.com/78769464/111478927-82d8e600-86fe-11eb-8dfc-e7e8dfa007cb.png">


After creating these user friendly options that provide valuable data analysis of their stock of interest in comparison to other stocks - I decided to refactor my set of data to allow for a quicker run time and a quicker feedback for Steve's parents. By doing this, I combined multiple steps. The most important change in the VBA is instead of having the macro run every single line of data - I changed it to run the tickerIndex of the data which combines the multiple rows of the same stock and allows the code to only have to run the one row of data. 

Since this is macro has combined commands and is more compact compared to the initial set of analysis - I added multiple lines of comments to give clear feed back of my process to anyone that would need to read the VBA to decipher the process. 

<img width="662" alt="4 Refactor" src="https://user-images.githubusercontent.com/78769464/111479945-7608c200-86ff-11eb-87fd-9c6507e3e936.png">

<img width="715" alt="5 Refactor" src="https://user-images.githubusercontent.com/78769464/111479953-786b1c00-86ff-11eb-9599-cfce3cef8f3f.png">

<img width="711" alt="6 Refactor" src="https://user-images.githubusercontent.com/78769464/111479972-7bfea300-86ff-11eb-88ba-13c4a4c9c1fa.png">

My end result was a quicker processing time, user friendly and an easily visualized outcome for all 12 stocks for Steve's parents.

![refactor time](https://user-images.githubusercontent.com/78769464/111570080-041d9080-8772-11eb-98f7-ca085f3d423c.jpg)


Advantages/Disadvantages to refactoring:

There are clear advantages to refactoring data such as minimizing run time and getting a result quicker but also cutting down the amount of code the computer has to delegate through. It also minimizes memory space and prevents any lags due to limited memory. It also gives a more compact version to completing it for the other people that want to read the macros you built. The disadvantage is the same situation to compacting it- it may be difficult to follow along with the code if you have an error and to easily pin point where that error is coming from, versus having everything singled out. 


Advantages and disadvantages of the original and refactored VBA script:  


The advantages to the original script is that it was more detailed and easier to follow along to someone that is not familiar with writing code or reading it. The main advantage to the refactored VBA script is that its quicker, takes less memory, and more compact. 
