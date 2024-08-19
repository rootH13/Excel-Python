I created this python code to be able to organize spreadsheets and data that come in the form of a csv.
I initially created this project in JupyterLab, and adjusted it to a .py format to run locally.

I used standard data science/organization python skills for dataframes such as .dropna, .head() and more to view all of my data.
My first step was to remove any null data that appears blank in Excel. I used .head(), and 'df' to view all of the data to ensure I did not accidentally delete data from a colum or row with almost all NULL values.

My first goal was to ensure that after sorting and cleaning up the csv, all of the data would be properly wrapped/formatted to each cell to avoid manual adjusting within excel.
I researched many StackExchange posts, and eventually came to use the tk python module to wrap the text. I consulted ChatGPT to clean-up my code and look for any alternatives in python modules.

Upon testing this script locally, it holds up to removing NULL columns and ensuring that data that is present in a column with almost all NULL values is not removed. All the data is properly wrapped and fitted to each cell, and data presented as "########" in excel is readable.

I then updated my script to be a PowerShell script (.ps1). It is now fully functional, with instruction message boxes as well (thanks to this StackOverflow post: https://stackoverflow.com/questions/2963263/how-can-i-create-a-simple-message-box-in-python).

To be able to run locally, if using PowerShell, type "python --version". If you receive an error, you must download and install python on your local device. If you already have Python installed, then you may proceed with running the script. Ensure your permissions are set appropriately to run scripts.

Here are the following modules you may need to install via 'pip install' to ensure this program runs:
  *tkinter
  *openpyxl
  *tabulate
  *easygui
