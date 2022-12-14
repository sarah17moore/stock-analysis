# VBA for Wall Street
A classwork example project reviewing stock portfolios using Excel and Visual Basic for Applications

---

# Overview
This is a classwork example in which "my friend" Steve is a recent finance graduate with his first client - his parents. Steve's parents have decided to invest money into DAQO New Energy Corporation. Steve wishes to diversify their portfolio, so he requested to look at data for several stocks in an automated and efficient manner. This analysis will allow Steve to look at summary statistics in 2017 and 2018 easily and with the quickest run time possible. 

## Purpose
We will use Visual Basic for Applications to write a macro process to help Steve compare average daily total volumes of several stocks at once for a chosen year. Essentially, the client believes that market forces will appropriately control the stock price and therefore wishes to know how actively the DAQO stock was traded within a year. Steve wishes to use the same reasoning to show his clients alternate stocks that can diversify the portfolio in order to reduce risk. 

Stock data can be gathered and presented for 12 different stocks in less than half a second with the click of a button and the entry of a year. Easy to read formatting will be added to the Excel sheet so that positive and negative returns can be visually identified. 

After an initial code was created, the code was refactored to speed up the macro process run time. 

---

# Results
The original code to complete this analysis did not fully utilize variable assignment. This is a screenshot of the original code showing that full equations were typed out in several steps:

![2018 - old, slower code example](/Resources/2018_old_code_screenshot.png)

The new and improved code utilized variables which helped the process repeat actions quicker. This is due to the macro not needing to read full equations or prompts repeatedly, instead, the equation can be read once and "stored" in the memory of the variable. This is a screenshot of the new and improved code that shows that variables have been utilized to accomplish the same actions as the old code:

![2018 - new, quicker code example](/Resources/2018_new_code_screenshot.png)

By eliminating several steps in the macro code, the run time was shortened from 0.273 seconds to 0.028 seconds. 

![2018 - old, slower code run time](/Resources/VBA_Challenge_2018-(original).png)

![2018 - new, quicker code run time](/Resources/VBA_Challenge_2018.png)

In short, DAQO was not as actively traded as Steve's clients may have hoped. The return has shown to be negative. Alternatively, stocks ENPH and RUN have positive returns. Steve's clients may be interested in investing into these stocks instead. 

# Summary
Refactoring code can be beneficial as this action can help keep process run times low, make Excel data sheets more efficient, and can help code be more readable by casual users or others who will need to edit the code in the future. On the other hand, if a original code works well and still completes the process in a reasonable amount of time, it may take extra time and effort to refactor the code. It may not be worth it to spend this time to refactor code that already works. 

These ideas can be applied to this refactoring process. While this refactored code is definitely easier to read and understand, which will help Steve if he needs to make adjustments or additions in the future, it did take a significant amount of turnaround time to finish the refactoring. Steve may have needed to decide if waiting for refactoring for ease of use in the future is worth it to him and his clients in the present.  
 
 ---
