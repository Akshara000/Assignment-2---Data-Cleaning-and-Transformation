# Assignment-2---Data-Cleaning-and-Transformation
Excel- Data Cleaning and Transformation 
Excel Assignment 2 - Data Cleaning and Transformation

I have completed Assignment 2 using Excel.

Solution:

1).Handling Missing Values:
• Check for missing values in the 'Price' column. How would you handle products with missing price information?
Answer:
Step 1 - Select the Range of price column and using Conditional Format Highlight the cells that contain blank. 3 Cells are found blank.
Step 2 - Find median . Here the median is 130
Step 3 - Replace highlighted blank Cells with 130

• If there are products with missing categories, propose a strategy to impute or deal with these missing values effectively.
Answer:
Step 1 - Highlight blank cells using conditional format on Category Column.Found that 4 cells are blank whose product name are - Backpack,Sneakers,Coffee maker,Fitness Tracker which are already categorised except Backpack.For the case of Backpack I guessed it will come under Accessories.
Step 2 - Create a new column name ‘Category2’ the used the formula =IF(F2<>"", F2, IF(B2="Sneakers","Fashion", IF(B2="Coffee Maker","Kitchen", IF(B2="Fitness Tracker","Electronics", IF(B2=“Backpack”,"Accessories","Uncategorised"))))).
Step 3 - Copy the cells of Category2 and paste the values on Category column.And Delete column Category2 .So that all blank cell in category impute with correct values.


2) Correcting Inconsistent Data:

   • Identify any inconsistent text formats present in the "Product Name" column.
Answer: Product name “laptop”, “smartphone” and “headphones” were lowercase 
 
  • Identify any typos present in the "Category" column.
Answer: “Electroni” appeared 5 times instead of Electronics

  • Use the find and replace function to standardize the text formats in the "Product Name" column and fix any typos or misspellings in the "Category" column.
Answer:Using Find and replace function replaced inconsistant(lower case) to Title case on product name column and replaced ‘Electroni’ to ‘Electronics’.

3) Removing Duplicates:

   • Identify any duplicate rows within the dataset based on the entirety of each row, and remove them if any.
Answer:
Step 1- Select the entire range of table 
Step 2- Use Remove Duplicates Option on the ‘Data’ Ribbon
4) Splitting and Merging Data:
  
Split the "Product ID" column into two separate columns for " Manufacturing Date" and "Country Code". Remove unnecessary characters, if any.

Answer : 
Step 1 - Create two Columns named  " Manufacturing Date" and "Country Code”
Step 2 -Type “28-JAN” on first cell under ‘ Manufacturing Date’ Column and “US” on ‘Country Code’ .Because the First cell of ‘Product ID’ is “28-JAN-US”. Here we given which pattern the two new Column Should perform.
Step 3- Ctrl+E (flash fill) applied on these columns.

5) Number Formatting:

Format the data type of the "Price" column to currency format. 

Answer:
Step 1 - Select the price Column.
Step 2 -Right Click then take the option Format Cells.
Step 3- From Format Cells pop up window Select Category as Currency format.

Format the "Manufacturing Date" column to display dates in the "DD-MM-YYYY " format. 

Answer:
Step 1 - Select ‘Manufacturing Date’  column.
Step 2 -Right Click then take the option Format Cells.
Step 3- From Format Cells pop up window Select Category as Date format and select type as ‘DD-MM-YYYY”


6) Conditional Formatting:
Apply data bar or color scales conditional formatting in the "Price" column.
Answer:
Step 1 - Select Price range
Step 2 - Select data bars option from conditional format.
Step 3- Choose a gradient

Create a custom rule for conditional formatting in the "Category" column to highlight cells where the category is “Electronics."
Answer:
Step 1 - Select “Category” range
Step 2 - Select New Rule option from conditional format.
Step 3 - Select Style as Classic and  Select ‘Use a Formula to determine which cells to format’
Step 4 - Use the Formula =$F2=“Electronics"
Step 5 - Format with Yellow fill with yellow text chosen to highlight.
