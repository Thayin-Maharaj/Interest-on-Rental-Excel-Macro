# Interest-on-Rental-Excel-Macro
This macro was created to be able to solve for the interest earned when renting out office equipment to other parties. The spreadsheet uses the total cost of equipment, number of periods for rental and the monthly rental figure that was agreed upon by both parties. The macro then iterates on the solution for interest using a basic numerical method.

# Important Note
The macros are designed to only perform a maximum of 1000 iterations to prevent any infinite looping, if the program reaches this limit and does not converge to 0 then it is recommended to set the interest value in the bold block to a higher value such as 100% then rerun the macro.
