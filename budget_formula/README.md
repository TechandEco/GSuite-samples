# Simple Google Apps Script custom function for budget planning

- Follow me on
  [Twitter](https://twitter.com/TechandEco)
- Visit this solution's step by step
  [instructions with screenshots here](https://medium.com/@TechandEco/faster-budgeting-with-a-google-apps-script-custom-function-2f1dfd037d2d)

## Introduction

- This is the 2nd of 5 Apps Script tutorials for users in Finance.

- In this sample I walk you through how to create your own
 [custom function](https://developers.google.com/apps-script/guides/sheets/functions)
 so you can prepare your annual household budget by entering your expenses in
 one row and marking their frequency in another (ex: annual, monthly, weekly,
 daily, or only one time). When you have rows and rows of expenses with
 different time periods, creating your own formula with a Google Apps Script
 helps you save a lot of time.

- This is because when you have all your budgetary information entered in your
  sheet, you can use it to auto calculate the total annual amount of each
  expense by entering the formula ex: `=ANNUAL_COST()`. This multiplies your
  expense by 52 weeks if you mark it as annual, or 12 months if marked
  monthly.

## Try it

1. [Copy this Google Sheet](https://docs.google.com/spreadsheets/d/126YIKW8NO_6GqyZNF1jbUZQIfb-rLwNyIH79qdeELJU/copy)

1. Organize columns and data as listed in
   [blog post](https://docs.google.com/spreadsheets/d/126YIKW8NO_6GqyZNF1jbUZQIfb-rLwNyIH79qdeELJU/copy)

1. Apply custom function by using `=ANNUAL_COST(D2, C2)`
