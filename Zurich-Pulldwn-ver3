#!/home/michael/PycharmProjects/Zurich-Linux-Pulldwn-ver3/.venv/bin/python3.10
import json
import subprocess
from urllib.request import urlopen

import pandas as pd
from openpyxl import load_workbook

from checkdate import check_date  # this is a function I created as a single file elsewhere in this project

# to check the date cell in the existing source file for comparison with the latest DF date column


""" This is a code version of a project originally set for a windows environment which has been tweaked for running on
linux OS version. So, path to files format has been changed to reflect that.
this version 3 has fund name changes to reflect fund mix changes in 2025 which means a new Excel sheet for 2025 also
I have also added a linux OS specific message box actions which use the "subprocess" module to write messages
to the screen advising of progress completion.
We are pulling down the API data in json format using an API link process by John Watson Rooney.
it is a link found when inspecting the requests made by the 
dynamic table in the target site so that we can get behind the dynamic table data and pull down the source data 
because dynamic data can be  erratic when being pulled down with selenium from some sites. 
We are using urlopen to get the site API json data then, we will read the json file and convert it to a dictionary 
before creating a pandas data frame object which, we can then finally write to an excel file using openpyxl. 
This version has updated fund row numbers to reflect the new fund mix agreed in 2024. In order to use
this script later on as an automated process running in the back ground, I introduced an auto check code to the API 
which compares the date on the Excel file we are going to write to with the date on the API data being
processed in the for loop. If they are the same then, we do not want the code to keep overwriting data to the
file every time we start the pc on the same day. If they are different then, the logic allows the rest of
the script to come into play and post the data to the excel file. The function is called "check_date"""

url = "https://www.zurichlife.ie/services/listFundPrices?searchSetupCompany=1&searchFundGroup=10&searchToDate="
response = urlopen(url)  # we save the response in this variable for later use
# we create a set of empty lists to hold our loop results data, so we can use it later in our pandas process
fund = []
date = []
bid = []
sell = []
data = json.loads(response.read())  # this is the option used with the Urlopen method to read the json response string
# then we look through the response string to find the dictionary with a key name that contains the funds we want
# the key name we want will have multiple data values after the :. So, all we then need to do is loop
# through the entire dictionary looking for particular headings we want and store the associated values in a list
temp_one = (data["fundPriceList"])  # this is the key word dictionary we want
for item in temp_one:
    fund_name = item["fundDesc"]
    fund.append(fund_name)
    price_Date = item["priceDate"]
    date.append(price_Date)
    bid_Price = item["bidPrice"]
    bid.append(bid_Price)
    sell_Price = item["offerPrice"]
    sell.append(sell_Price)
    # each time we loop through, we add the result to our previously created empty list
    # now we create a dictionary with key Headings and the resultant values posted to each list for use in pandas
    result = {"Fund": fund, "Price date": date, "Bid": bid, "Offer": sell}
# now we run a comparison with the output from check_date function I created with the date we have got from
# the processed API "price date" data(above) to make sure the dates are not the same. If they are, the process stops
# after the "if" statement.
value = check_date()
sheet_date = date[0] # We check the date value for the first entry in our dictionary Key value "date" list
if value != sheet_date:
    # noinspection PyUnboundLocalVariable
    df = pd.DataFrame(result)
    print(df)
    df.to_csv('mytest-headless.csv')  # this is a check file, so we can  see how the first DF contents looked like
    # in case we need to trouble shoot later if website rows are changed by creator
    # please note the file name came from a previous demo project I created from a selenium project and just
    # pasted straight in here. it relates to headless mode in selenium but means nothing in this project.
    # now we have a pandas data frame we can filter and export into Excel later
    # next thing to do is filter the data frame by line index number to geta a new data frame with funds we want
    # we do this using the pandas df.loc argument to produce a second data frame using certain rows from the first
    # data frame which has 64 rows of data. we do not want all of them.
    df2 = (df.loc[[4, 32, 29, 42, 48, 61, 59, 40, 21, 25], ['Fund', 'Price date', 'Bid', 'Offer']])
    # the new data frame df.loc method above has the required rows listed first followed by the headings we want
    # following the change in the fund mix in 2025
    """please be aware that  Zurich move the fund rows around from time to time as and when they change funds
    So if you run this df.loc filter and an error comes back saying the key index was not present it maybe
    because the new Zurich source file has had rows added or subtracted and the row index being looked
    for does not exist any more. In which case you must edit the row index selection in the df.loc filter list."""
    """ Note:, the values in column 'Bid' and 'Offer' will appear in the target excel sheet as string objects and not
    floating integer numeric values. This will mean Excel will not be able to work with them if there is a 
    calculation happening in the sheet using these data type values. So,we change the data type of the values in
    the selected column of our df2 data frame with the Pandas .astype() function (see next two lines)"""
    df2['Bid'] = df2['Bid'].astype(float)
    df2['Offer'] = df2['Offer'].astype(float)
    """ now we start the process of opening the excel work book we will copy the data frame into
    """
    FilePath = "/home/michael/Desktop/new fund mix anlysis Zurich25.xlsx"
    # Generating workbook instance with our existing Excel file above as the target (this was chnaged to the 2025 version for ver 3)
    ExcelWorkbook = load_workbook(FilePath, keep_vba=True)
    # Generating the pandas writer engine to copy the new data frame to the new sheet tab
    with pd.ExcelWriter(FilePath, engine='openpyxl', mode="a", if_sheet_exists="overlay") as writer:
        df2.to_excel(writer, sheet_name='analysisnew', startcol=0, startrow=1, index=False)
    subprocess.Popen(['notify-send', "Sheet data updated with latest data"])
else:
    subprocess.Popen(['notify-send',"Sheet dates matched in API so process stopped"])
#   the "else" statement above causes a message to appear if the dates match with data already on the Excel sheet

""" so the important thing in the "write to" action above, is first to add the "a" attribute to the Excel-writer. 
This tells it to only append data and not over write other data already in a file. 
As I want to post data to a target sheet in the existing file. 
I use sheet_name attribute to tell excel writer the name of the sheet. In this case the sheet already exists.
this sheet will be my target sheet. Finally, as i want the program to write data to this target sheet any time I run it, 
I have to tell it what to do when it finds the target sheet already existing in the target file. 
I do this by  adding the attribute "if_sheet_exists"to the excel writer set up. I tell it to "overlay" the sheet 
with the new data it is trying to write to the file. Otherwise the programme would crash with an error saying 
it had found the sheet already exists. finally, note that I have told it what column and row to start overlaying 
the data to. This is done with the arguments 'startcol=' and 'startrow=' in the ExcelWriter action.
"""
