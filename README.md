# FitRowXL
 Dynamically Autofit Row Height with VBA UDF Function
 
(I am Vietnamese, so I translated the article into English)

*** LATEST UPDATE: 10:40 19/4/2021 ***

Today, I will share with you a VBA UDF function that automatically stretches lines when the value in the worksheet changes, the line stretching will also change automatically.
The code below is the first version, so there may be certain errors, so you should check first, if the code is stable, you can use it.

I decided to write this code because there are many posts on the forum that ask about automatic line spacing, and when I read through those posts, no one can handle the problem thoroughly, or the code is not dark. pros, or code that does not handle multiple regions that have the same line.

And the UDF function is the best way to help you do not have to re-code, but just write the function as a normal Excel function to perform line stretching.



UDF Function S_FitRow:

### Function:
1. Fit Row automatic.
2. Include merged cells.
3. Stretch the line with multi-cell values on the same line.
4. Include Print mode.
5. Add a certain height after line reduction.
6. Because of using UDF function, it is very optimal and saves CPU.

### Instructions for using the function:

Order	Parameters	Type	Ability
1	Target	The area to be stretched	Get the area to be elastic
2	Margin	Number type	Increase the height by some
3	defaultHeight	Number type	The default height if the value is empty
4	IncludeNoWrap	Yes / No	Elasticity even with zero WapText
5	Title	Chain	Any string set by the user (Otherwise, return the value to Fit: {area})

How to write the function quickly, type in the string =S_FitRow and press Ctrl + Shift + A.

### Example: = S_FitRow (A1: E9, 5, 40, FALSE)
+ A1: E9 is the area to be scaled
+ 5 is the height increase by 5 units
+ 40 is the default height if all values are empty
+ FALSE then the cell is not WapText is inelastic.
You can also type quickly =S_FitRow(A1: E9), ignores the settings

### Note: the function only does line stretch if the WrapText mode of the cells is active. Add-in Methods:
1. Type S_FitRow_OFF function: If you are editing a worksheet, turn off line stretching or turn on Design Mode in the Developer Tab.
2. Type S_FitRow_ON function: If auto line stretching is turned off, turn it on.
3. Procedure S_FitRow_Toggle + Check box named chxAutoFitRow is used to turn on and off line spacing if desired (The example is in Sheet1 in the attached file below on Github).
Step 3 is a procedure to prevent application code calculated at just kicked off, as may encounter status code will slow down the boot process.
Let the following code in the Workbook_Open event: Call S_FitRow_Off
Please reopen in step 2 or step 3.

*** If you write too much formula, the indispensable function S_FitRow this step 3

How to enter multiple arrays:
Method 1: Use the Excel Indirect function: = S_FitRow (INDIRECT ({"A1: C9", "D2: D3", "E5: E6"}), 5, 40)
Method 2: Use S_Cells user-define function: = S_FitRow (S_Cells (A1: C9, D2: D3, E5: E6), 5, 40)


### Utilize the function:
With 2 following examples:
1. How to write the general function as follows will cause slow and CPU consumption:
= S_FitRow (A1: Z500)
2. Writing a single function for each region saves, and the code will run faster:
= S_FitRow (A1: Z1)
= S_FitRow (A2: Z2)
= S_FitRow (A3: Z4)

That is, enter the area so that the function corresponds to the largest number of merged rows of the corresponding row.
If we have the combined regions A1: C9, D2: D3, E5: E6, then it is clear that A1: C9 has lines containing all the remaining regions,
So we write: = S_FitRow (A1: E9)

Possible errors:
Because the code will borrow a cell in the worksheet as the cell for line spacing, an error occurs if the function has a reference range intersecting the borrowed cell.
If there are two functions that refer to the intersection of the two regions, an error may occur.
If the sheet or cell reference is locked, an error may also occur.

*** Please note that the code may not be optimal, so it can be updated many times, so if you use the code, you should regularly review the article, there will be an update notice if any. at the beginning of the post.
