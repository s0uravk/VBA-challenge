#VBA-challenge

![VBA](https://github.com/s0uravk/VBA-challenge/assets/144293972/208adcdd-61a9-4140-8987-58bf36e4e652)


Introduction : StockData Program will summarize essential parts of stockmarket such as Changes in stock price over the year, representated in Percentage as well. Then it shows which stock had gained the most, lost the most and been traded the most. And the best part is that it works through all the Worksheets present in a Workbook.

Requirements : It requires Data to be processsed to be formatted in a specific order. Ticker symbols being in the first column(A), Open price in the thrid column(C) and closing price in the sixth column(F) and volume in the seventh column(G), all this data sorted based on the date that observation is from i.e. ascending order.

Functionality : This Algorithm will create a column of summarized ticker symbols in ninth column(I) with difference between closing price on last oberservation and opening price of first observation of that specific ticker being in tenth column (J) and percetage of change occured in eleventh column (K) and sum of total volume of that ticker in twelfth column (L). And the columns with Yearly and Percentage change will be highlighted based on increase or decrease in prices.

Moreover, it will calculate the greatest percentage of increase , decrease in stock price as well as greatest total volume in cells Q2,Q3 and Q4 respectively and there ticker symbols being in the cells, P2,P3 and P4 accordingly. The last but not least feature is that, it will summarize not only one but all the Worksheets available in a Workbook.

Modification : A button can also be added by going to developer tab, selecting Insert in Controls and selection Button from Form Controls and then assign the name of project to that button. So all the data can be processed with just one click of a Button.

Refrences : Code snippet used to perform conditional formatting as per the MS excel documentation (https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions). Firstly, I used if-else statement to compare the value and then assigning the interior color index based on the value but using conditional formating can make the code more efficient and can be used outside the loop as well. 
For greatest gain and greatest loss, i used Worksheet Function Max and Min ,(https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.max) and (https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.min). And to find the index of greatestgain, greatest loss and greatest total volume, i used Worksheet Function Match (https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.match), That later retrieved the corressponding Ticker all those Values.
