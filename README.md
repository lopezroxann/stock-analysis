# Stock Analysis and Refactoring Code

## Overview of Project

### Purpose

We are assisting Steve in his analysis of the stock market to conduct better research for his parents.  Although we initially started off with a code for a dozen stocks, we will be refactoring the code to expand to the entire stock market.

## Results

With the original macro, we can see below that this took our system approximately 0.28125 seconds for each 2017 and 2018 to run the code:
<br/>

<table>
  <tr>
    <td>2017 Original Macro</td>
     <td>2018 Original Macro</td>
  </tr>
  <tr>
    <td valign="top"><img src="https://github.com/lopezroxann/stock-analysis/blob/main/Resources/2017_Original_Macro.png"</td>
    <td valign="top"><img src="https://github.com/lopezroxann/stock-analysis/blob/main/Resources/2018_Original_Macro.png"</td>
  </tr>
 </table>
<br/>

In comparison, once we refactored our code, our elapsed time improved to 0.08984375 and 0.0859375 seconds for 2017 and 2018 respectively:
<br/>

<table>
  <tr>
    <td>2017 Refactored Macro</td>
     <td>2018 Refactored Macro</td>
  </tr>
  <tr>
    <td valign="top"><img src="https://github.com/lopezroxann/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png"</td>
    <td valign="top"><img src="https://github.com/lopezroxann/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png"</td>
  </tr>
 </table>
<br/>

Looking at the changes to our code, the application of the tickerIndex has allowed us to loop through all the tickers simultaneously versus one by one. Also, removing the nested For loop in our original code in combination with the tickerIndex in the refactored code has reduced redundancies and permits the code to run through more quickly and efficiently. 

![alt text]( https://github.com/lopezroxann/stock-analysis/blob/main/Resources/Comparison_of_VBA_code.png)
<br/>
## Summary

- What are the advantages and disadvantages of refactoring a code?<br/>
Refactoring code has many advantages. It reduces the elapsed time of your system run through the code as the code is more concise. Having code that is well organised and clean permits others to follow and understand your code more easily. Also, refactored code takes up less memory and time as you are not writing a new code from scratch, but improving a current exisiting code. Although refactoring can also help us find bugs, a disadvantage is that you can also create bugs in the process. Which can in turn take more time to fix and cost more money.


- How do these pros and cons apply to refactoring the original VBA script?<br/>
As we can see in the image previously posted, the refactored code is more clean, concise, and easy to follow. There is evidence that the new refactored code ran significaly faster by adding the tickerIndex and removing the nested For loop. In the process of refactoring this code, I did come across a couple of bugs that were created mostly by syntax errors in my If-Then statements that were corrected by trial and error.

