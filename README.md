# AutoMailer:

## Summary:


## 

## Detailed Explanation:




Uses Python to obtain an access token from Outlooks API. Once requested and received began to extract data in the form of JSON from the Outlook credentials specified. This script can be ututlized for multiple Outlooks functionally but in this case it is used to check against a Data Integrity Agent for fraud emails. 

It allows data to be parsed out of the JSON and either complete a simple comparison or a direct comparison depending on which one is needed. The script can be customized to suite the need for the user but as of right now it allows the given data to be stored in the form of .txt files held on either a GitHub account to be pulled or directly onto a server. It begins by pulling the custom parsed data defined and dumps them into a particular folder. 

Once that process is complete the code begins to pull each piece of data to see if it links to a matched file. If there is no match then the data stays in the pulled data folder and continues on to the next. If the data does pertain a match then this is recorded and the data is deleted from the pulled data file. This continues until all the data is checked. At the end of the script the fraud emails can be detected by checking the data that has not been erased in the pulled data folder. Of course there is much more to extend off of this project but the only necessary goals required this process. For instance, the script could the continue to give certain information about the error in comparison checks for each case. Along with this there could also be functionality to speed up the process of comparing each email which can either involve a binary search.

