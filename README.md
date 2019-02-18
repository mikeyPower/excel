# work_stuff

## How to Count all errors in a range

1. In a blank cell, type this formula =SUM(IF(ISERROR(A1:C10),1)).

![alt text](https://raw.githubusercontent.com/mikeyPower/excelsior/branch/path/to/img.png)

2. Then press Ctrl+Shift+Enter keys together, and you will get the number of all the error values of the range.

![alt text](https://raw.githubusercontent.com/mikeyPower/excelsior/branch/path/to/img.png)


Note: In the above formula, A1:C10 is the range that you want to use, you can change it as you need.


## Count The Number Of Specific Types Of Errors In A Range


1. In a blank cell, please type this formula =COUNTIF(A1:C10,"#DIV/0!"), see screenshot:

![alt text](https://raw.githubusercontent.com/mikeyPower/excelsior/branch/path/to/img.png)

2. Then press Enter key, and the number of #DIV/0! error cells will be counted.

![alt text](https://raw.githubusercontent.com/mikeyPower/excelsior/branch/path/to/img.png)

Note: In the above formula, A1:C10 is the range that you want to use, and #DIV/0! is the type error that you want to count, you can replace it as your need.

## Count The Number Of Cells Ignoring Errors In A Range

If you want to count the number of cells without errors, you can use this array formula: =SUM(IF( NOT( ISERROR(A1:C10)),1 )), and then press Ctrl+Shift+Enter keys simultaneously. And all the cells ignoring error cells will be calculated (including blank cells). See screenshots

![alt text](https://raw.githubusercontent.com/mikeyPower/excelsior/branch/path/to/img.png)

![alt text](https://raw.githubusercontent.com/mikeyPower/excelsior/branch/path/to/img.png)
