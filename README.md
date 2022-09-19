# VBA for Wall Street
A classwork example project reviewing stock portfolios

---

# Overview
This is a classwork example in which "my friend" Steve is a recent finance graduate with his first client - his parents. Steve's parents have decided to invest money into DAQO New Energy Corporation. Steve wishes to diversify their portfolio, so he requested to look at data for several stocks in an automated and efficient manner. This analysis will allow Steve to look at summary statistics in 2017 and 2018 easily and with the quickest run time possible. 

## Purpose
We will use Visual Basic for Applications to write a macro process to help Steve compare average daily total volumes of several stocks at once for a chosen year. Essentially, the client believes that market forces will appropriately control the stock price and therefore wishes to know how actively the DAQO stock was traded within a year. Steve wishes to use the same reasoning to show his clients alternate stocks that can diversify the portfolio in order to reduce risk. 

Stock data can be gathered and presented for 12 different stocks in less than half a second with the click of a button and the entry of a year. Easy to read formatting will be added to the Excel sheet so that positive and negative returns can be visually identified. 

After an initial code was created, the code was refactored to speed up the macro process run time. 

---

# Results
using images and examples of code, compare the stock performance between 2018 and 2018 as well as the execution times of the original script and the refactored script

The original code to complete this analysis did not fully utilize variable assignment. This is a screenshot of the original code showing that full equations were typed out in several steps:

The new and improved code utilized variables which helped the process repeat actions quicker. This is due to the macro not needing to read full equations or prompts repeatedly, instead, the equation can be read once and "stored" in the memory of the variable. This is a screenshot of the new and improved code that shows that variables have been utilized to accomplish the same actions as the old code:

By eliminating several steps in the macro code, the run time was shortened from 0.273 seconds to 0.028 seconds. 
![2018 - old, slower code run time](/Resources/VBA_Challenge_2018.png)

![2018 - new, quicker code run time](/Resources/VBA_Challenge_2018-2.png)

In short, DAQO was not as actively traded as Steve's clients may have hoped. The return has shown to be negative. Alternatively, stocks ENPH and RUN have positive returns. Steve's clients may be interested in investing into these stocks instead. 

# Summary
in a summary statement address:
 What are advantages or disadvantages of refactoring code
 how do these pros and cons apply to refactoring the original vba script
 
 ---

there is a title, multiple paragraphs each with headings, subheadings to break up text, links are working and images are formatted/ displayed where appropriate
 
 
