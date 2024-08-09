I created this python code to be able to organize spreadsheets and data that come in the forms of csv.
I initially created this project in JupyterLab, and adjusted it to a .py format to run locally.

I used standard data science/organization python skills for dataframes such as .dropna, .head() and more to view all of my data.
My first step was to remove any null data that appears blank in Excel. I used .head(), and 'df' to view all of the data to ensure I did not accidentally delete data from a colum or row with almost all NULL values.

My first goal was to ensure that after sorting and cleaning up the csv, all of the data would be properly wrapped/formatted to each cell to avoid manual adjusting within excel.
I researched many StackExchange posts, and eventually came to use the tk python module to wrap the text.
