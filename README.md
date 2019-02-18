# work_stuff

## How to Count all errors in a range

1. In a blank cell, type this formula =SUM(IF(ISERROR(A1:C10),1)).

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/1.png)

2. Then press Ctrl+Shift+Enter keys together, and you will get the number of all the error values of the range.

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/2.png)


Note: In the above formula, A1:C10 is the range that you want to use, you can change it as you need.


## Count The Number Of Specific Types Of Errors In A Range


1. In a blank cell, please type this formula =COUNTIF(A1:C10,"#DIV/0!"), see screenshot:

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/doc-count-errors3.png)

2. Then press Enter key, and the number of #DIV/0! error cells will be counted.

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/5.png)

Note: In the above formula, A1:C10 is the range that you want to use, and #DIV/0! is the type error that you want to count, you can replace it as your need.

## Count The Number Of Cells Ignoring Errors In A Range

If you want to count the number of cells without errors, you can use this array formula: =SUM(IF( NOT( ISERROR(A1:C10)),1 )), and then press Ctrl+Shift+Enter keys simultaneously. And all the cells ignoring error cells will be calculated (including blank cells). See screenshots

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/6.png)

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/7png.png)


## Highlight rows based on a certain criteria

Suppose you have a dataset as shown below and you want to highlight all the records where the Sales Rep name is Bob.

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/highlight/example-table.png)

Here are the steps to do this:

1. Select the entire dataset (A2:A17 in this example).
2. Click the Home tab.

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/highlight/Home-Tab-in-the-Excel-Ribbon.png)

3. In the Styles group, click on Conditional Formatting.


![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/highlight/Click-on-Conditional-Formatting.png)

4. Click on ‘New Rules’.

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/highlight/Click-on-New-Rule-Highlight-Rows-Based-on-a-Cell-Value-in-Excel-Conditional-Formatting.png)

5. In the ‘New Formatting Rule’ dialog box, click on ‘Use a formula to determine which cells to format’.

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/highlight/Use-Formula-Option-to-Highlight-Rows-based-on-cell-value.png)

6. In the formula field, enter the following formula: =$C2=”Bob”

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/highlight/Specify-the-Formula-to-Highlight-rows-if-it-is-True.png)

7. Click the ‘Format’ button.

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/highlight/Click-on-the-Format-Button.png)

8. In the dialog box that opens, set the color in which you want the row to get highlighted.

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/highlight/Color-to-Fill-to-highlight-the-rows.png)

Click OK.

This will highlight all the rows where the name of the Sales Rep is ‘Bob’.

![alt text](https://github.com/mikeyPower/work_stuff/blob/master/Images%20for%20excel/highlight/All-rows-where-name-is-Bob-are-highlighted.png)

Note that the trick here is to use a dollar sign ($) before the column alphabet ($C1). By doing this, we have locked the column to always be C. So even when cell A2 is being checked for the formula, it will check C2, and when A3 is checked for the condition, it will check C3.

This allows us to highlight the entire row by conditional formatting.
