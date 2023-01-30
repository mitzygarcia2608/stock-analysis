# Green Stocks Analysis


# Overview of the Project

Steve is a recent finance degree graduate, and his parents are looking to invest their money into DAQO New Energy Corporation.  To tackle this task, a VBA code was written to analyze the stocks.
However, the dataset now has been expanded to include the entire stock market over the last few years.  The task is to refactor the current code to analyze a larger amount of data, while reducing the execution time for running the code.  The code will loop through all of the data through VBA.

# Results
To reduce the execution time, one of the main adjustments made was to initialize the tickerIndex and increase the value of the variable by 1 when the ticker name changed.  The code then loops through all of the arrays to output the name, volume, and return for each defined ticker.
![image](https://user-images.githubusercontent.com/111592990/215458619-70ed004f-adb1-402c-afef-17ee43e03462.png)

As you can you see in the screenshots below, the refactored code runs faster than the original code, while outputting the same results:
![image](https://user-images.githubusercontent.com/111592990/215458692-1747fd49-67b1-485e-a84d-759aad061446.png)
![image](https://user-images.githubusercontent.com/111592990/215458728-0b28016b-0e6a-478b-8d08-1dbb70d1853d.png)
![image](https://user-images.githubusercontent.com/111592990/215458770-8101701a-ed4f-4de1-a18a-f51d73b08821.png)
![image](https://user-images.githubusercontent.com/111592990/215458791-74447403-4a70-46d0-891f-6b0e5ce1350c.png)
# Summary
The advantages of refactoring code is to make the code run quickly and efficiently.  By having a refactored code, you can make it easier for other users to use and share your code.  For example, if you created a code for a specific project, and another user in the future needs to update your file, the cleaned up version makes it easier to edit and understand.
One of the main disadvantages of refactored code is the level of understanding and complexity it may take to achieve the correct result.  Since the code will most likely be more complex than the original, it would take much more time to write, debug, and finalize a compound refactored code.
For the original code in the project, the biggest disadvantage was the run time of the code.  It only had one ticker variable, which caused the code to run through the worksheet multiple times for each different stock.  The refactored code incorporated the tickerIndex variable to track the different stocks as the code ran, which helped decrease the overall run time of the code.
