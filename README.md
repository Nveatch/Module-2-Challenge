# VBA Stock Analysis

## Overview of Project

### Purpose
The purpose of this challenge is to take the original VBA all stock analysis, and refactor it. We're doing this in order to make the macro more efficient by removing the requirement of having to loop through the entire table multiple times. This allows the macro to be used with much larger data sets without taking excessively long to complete.

## Results

![Original vs. Refactored Run Time](/resources/Side by Side Time Comparison.png)
From the attached picture, we can see the run time was reduced from about **0.60 seconds** to **0.08 seconds** (a decrease of **86.7%**) for 2017, and from about **0.61 seconds** to **0.08 seconds** (a decrease of **86.8%**) for 2018)

![2017 vs. 2018 Annual Return Percentages](/resources/Annual Return Graph.png)
From the attached graph, we can see that 2017 was a good year for the market, with most of the stocks having a positive return, while 2018 was a bad year for the market (mostly negative returns). From this comparison, I would conclude that good tickers to invest in are "ENPH" and "RUN", as they had a net positive return both years. Of those two, "ENPH" is more consistent, but both had similar return percentages for 2018.

Conversely, stocks to avoid would be the tickers "TERP", "AY", and "CSIQ". "TERP" had negative returns both years, and both "AY" and "CSIQ" essentially broke even over the two years (with "CSIQ" fairing slightly better).  

## Summary

### What are the advantages or disadvantages of refactoring code?

The advantages of refactoring code are an increase in efficency/elegance, and just a general improvement in understanding how the code works. By going line by line through the original code during the refactoring process, you get a deeper understanding of how every section works, and how the code performs its function as a whole. After refactoring, not only is the code more efficient than the original, it's often cleaner as well (and hopefully better notated). In addition, refactoring means you don't have to start from scratch, as it allows you to take snippets from the original code. 

The disadvantage of refactoring code is that you have to sift through potentially someone elses code, in order to understand it well enough to refactor it. Not only could this be frustrating if the code isn't well notated, but it could also be more time consuming than writing new code from scratch.

### How do these pros and cons apply to refactoring the original VBA script?

In regards to refactoring this script, the advantage of the code being more efficient definitely applied, as the run time was significantly shorter after changing the code. Also, while the changes from the original to the refactored code were fairly significant in my opinion, there was some time saved due to being able to copy/paste snippets, and I understand the original code better now after thinking through this challenge.

The disadvantage I listed doesn't really apply to this refactoring, as the original code was my own, and due to using the comment lines from the module, was easy enough to follow and break down and refactor. 