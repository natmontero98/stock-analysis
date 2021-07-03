# Stock Analysis

Stock analysis in Excel (VBA), Module 2.

## Overview of the project

This analysis aims to help Steve, a recent finance graduate, perform a stock analysis. Steve wants to analyze the green energy stocks because his parents are passionate about green energy. They have already invested all their funds in one stock (DQ), but Steve is concerned about diversifying their funds, so we want our analysis to tell us which are the potential best stocks to invest in. 

For this analysis we'll be using VBA, a programming language that interacts with Excel. 

The results we want to get for each stock are the total daily volume (total number of shares traded throughout the day) and their yearly return.

It's important to mention that our first code ran well for our 12 stocks, but Steve wants to expand the dataset to include the entire stock market over the last few years, so we will refactor the code to make it more efficient.

## Results

In both codes we achieved the same results for total volume and returns of the stocks. The final results were, for each year, the following:

<img width="206" alt="results 2017" src="https://user-images.githubusercontent.com/85467925/124342318-da00aa80-db77-11eb-835d-f5d9ad1350ff.png">
<img width="205" alt="results 2018" src="https://user-images.githubusercontent.com/85467925/124342322-e08f2200-db77-11eb-95fd-c581f06b8a21.png">

Thanks to contidional formatting we can tell that most stocks performed better in 2017.

### First version of the code

Our first code where we used multiple variables, only one array and nested for loops, had the following running times for each year analyzed:

* 2017: 0.359375 seconds
<img width="313" alt="2017 initial time" src="https://user-images.githubusercontent.com/85467925/124342405-99556100-db78-11eb-8cb0-36f554b8ceaf.png">


* 2018: 0.382812 seconds
<img width="316" alt="2018 initial time" src="https://user-images.githubusercontent.com/85467925/124342407-9ce8e800-db78-11eb-97ee-c86ff594f10e.png">


### Refactored code

This code loops through the data one time to collect all of the information. After refactoring, the running times were the following:

* 2017: 0.070312 seconds
<img width="314" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85467925/124342352-2fd55280-db78-11eb-8a6a-9d984f31cf94.png">

* 2018: 0.078125 seconds
<img width="307" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/85467925/124342360-3a8fe780-db78-11eb-9348-1197ede60e0e.png">



We can tell that the refactored code had lower running times that the first one, so the objective was met.

## Summary

### Advantages and disadvantages of refactoring code

The main advantage of refactoring code is that it allows you to perform the same functions and achieve the same results in a lower amount of time. It makes code more efficient, easier to understand, and more organized. This is especially beneficial when we're working with very large databases. 

A disadvantage, on the other hand, is that refactoring might not be an easy task, especially when the code is large and complex. Refactoring requires resources like time and money so it's important to evaluate in each case if the advantages outweigh the disadvantages.

### How do these pros and cons apply to refactoring the original VBA script?

In this analysis, the code itself was not very large or complex, so refactoring was possible in only a couple of hours. Now that our code is shorter, cleaner, and runs faster, we can analyze the whole stock market as Steve needed. So we can conclude that, in this case, it was worth taking the time to refactor the code.
