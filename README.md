#   Microsoft-Excel-Project-5

# Case study

You are helping a colleague, Lucas, to create an Excel worksheet that tracks the sales results for one of Adventure Works' most popular products, the A2Mountain Bike Frame. Lucas must present the results during Adventure Works’ monthly sales review meeting. You need to add formulas to the worksheet to complete this report for Lucas.

⦁ The total revenue from the A2 Mountain Bike Frame sales for April.

⦁ The number of frames sold.

⦁ The lowest and highest daily sales figures.

⦁ The number of days in the month.

⦁ And the overall daily sales average.
_____________________________________________________________________________________________________________________________________________________________________________________________________________________

# Preparing a Monthly Sales Report

# Overview

In this exercise, you were tasked with applying core Excel functions—SUM, AVERAGE, COUNT, MAX, and MIN—to analyze daily sales data for the A2 Mountain Bike Frames in April. The objective was to calculate key sales figures including total revenue, number of units sold, highest and lowest sales figures, and the overall daily average.

This reading provides you with a step-by-step guide on how these results were achieved, along with the formulas used. It also highlights techniques like using the AutoSum shortcut and Insert Function to simplify the formula creation process.
_____________________________________________________________________________________________________________________________________________________________________________________________________________________
# Steps Performed

# 1. File Setup

    ⦁ File Downloaded: Monthly sales report.xlsx
    ⦁ Worksheet: A2 Mountain Bike Frames
    ⦁ The worksheet contained daily sales data for April, and the task was to fill in cells C35 to C40 with the appropriate formulas.
    ⦁ The General Format was applied to cells C35 to C40.
    
# 2. Formulas and Calculations

The following formulas were created to calculate key sales metrics:

# a. Total Revenue (Cell C35)
    ⦁ Formula: =SUM(E4:E33)
    ⦁ Purpose: Calculate the total revenue for April from the sales data in column E.
    ⦁ Result: $23,059,600

# b. Total Units Sold (Cell C36)
    ⦁ Formula: =SUM(C4:C33)
    ⦁ Purpose: Calculate the total number of A2 Mountain Bike Frames sold in April.
    ⦁ Result: 115,298

# c. Lowest Sales Day (Cell C37)
    ⦁ Formula: =MIN(C4:C33)
    ⦁ Purpose: Identify the day with the lowest number of units sold.
    ⦁ Result: 2,560 units sold on April 30, 2023
    ⦁ Note: The date of the lowest sales day was manually typed in D37.

# d. Highest Sales Day (Cell C38)
    ⦁ Formula: =MAX(C4:C33)
    ⦁ Purpose: Identify the day with the highest number of units sold.
    ⦁ Result: 4,921 units sold on April 16, 2023
    ⦁ Note: The date of the highest sales day was manually typed in D38.

# e. Number of Days in April (Cell C39)
    ⦁ Formula: =COUNT(B4:B33) or =COUNTA(B4:B33)
    ⦁ Purpose: Calculate the number of days in April.
    ⦁ Result: 30 days
    ⦁ Note: Both COUNT and COUNTA functions worked since the dates were stored as numeric values.

# f. Average Daily Sales (Cell C40)
    ⦁ Formula: =AVERAGE(E4:E33)
    ⦁ Purpose: Calculate the average daily sales in dollars.
    ⦁ Result: $768,653
    ⦁ Note: The result was automatically formatted as Accounting.

# 3. Final Formatting
    ⦁ After entering the formulas, the Autofill feature was used to copy formulas down to other rows when needed.
    ⦁ The final calculations for Total Revenue, Total Units Sold, and other metrics were verified to ensure accuracy.
_____________________________________________________________________________________________________________________________________________________________________________________________________________________
# Key Concepts & Excel Techniques Used

    ⦁ SUM: For calculating total revenue and total units sold.
    ⦁ MIN/MAX: To identify the lowest and highest sales days.
    ⦁ COUNT/COUNTA: To determine the number of days in the month.
    ⦁ AVERAGE: For calculating the average daily sales.
    ⦁ Autofill: Used to copy formulas down to subsequent rows for faster calculation across the dataset.
_____________________________________________________________________________________________________________________________________________________________________________________________________________________
# Conclusion
With all required calculations completed, the Monthly Sales Report is now ready for review and presentation at the sales meeting. The formulas have been successfully applied to extract valuable sales insights for April.
_____________________________________________________________________________________________________________________________________________________________________________________________________________________
