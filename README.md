# Stock Analysis

## Overview of Project
Analysis of 2017 and 2018 share price data for 12 different companies to determine annual trading volume and returns to determine optimal investment strategy

## Results
The analysis shows a steep did in share price performance between 2017 and 2018. While in 2017, all but one of the companies' shares delivered positive returns, in 2018 all but two shares delivered losses. Both losses and gains were rather significant for the selected stock. None of the shares performed continously well, allowing for no clear investment hypothesis without further qualitative analysis.

### 2017 share price performance

<img src="https://user-images.githubusercontent.com/90064437/139564461-a0125d7e-0ae4-4aa6-bbb6-9a46658ffdff.png" width="227">

### 2018 share price performance

<img src="https://user-images.githubusercontent.com/90064437/139564524-052ef112-e289-43b5-b103-6c53d8afb62f.png" width="227">

## Refactored code performance

Refactoring the original code allowed us to significantly reduce the time it took to run the calculations. Using arrays to store individual results allowed the program to not calculate individual outputs by running along all the lines of code. Run time for both the 2017 and 2018 data was reduced by 80% when using the refactored code.

### Refactored performance increase for 2017 data

<p float="left">
  <img src="https://user-images.githubusercontent.com/90064437/139564694-55d0f206-eb97-4fd6-bdb7-36a1656ba60e.png" width="300" />
  <img src="https://user-images.githubusercontent.com/90064437/139564720-2fc5ac99-7454-4e98-87ab-65dfad390916.png" width="300" /> 
</p>

### Refactored performance increase for 2018 data

<p float="left">
  <img src="https://user-images.githubusercontent.com/90064437/139564763-8251029b-0f41-47b7-b0b4-3693537931cb.png" width="300" />
  <img src="https://user-images.githubusercontent.com/90064437/139564775-dd5db574-8639-4348-9256-4ef28c2c4df5.png" width="300" /> 
</p>

### Summary

*What are the advantages or disadvantages of refactoring code?*

Refactoring the code allows the program to perform better and utilize less memory space. It also aids in making the code look cleaner and more simple, which may allow others who review your code or collaborate on it with you to quickly understand what you have written and what your code does.

The disadvantage is that refactoring can introduce new bugs into the code, no longer allowing it to run properly. If one person refactors the code to a degree that it is no longer easily understandable to a second person looking at the code it can also create additional burdens and delay programming. This especially applies as the code gets more complex or needs to be continously changed or updated.

*How do these pros and cons apply to refactoring the original VBA script?*

In our code, refactoring led to a clear improvement in performace and made the code look cleaner and more simple by getting rid of the nested For loop.

On the other hand, the new script and its reliance on multiple arrays meant that we had to be extremely diligent in making sure that each array was applied properly and accurately connected to other arrays. Making one mistake with a variable or an array meant that the whole code could no longer run, and debugging it was harder.

