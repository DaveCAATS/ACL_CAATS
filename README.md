ACL_CAATS
=========
Purpose
The ACL script and related Visual Basic script will extract Outlook information including folder names and email metadata for fraud or other analysis purposes.  The script also extracts other items including appointments and could easily be modified to obtain additional information, including the contents of the email.
 
Approach
The application uses a Visual Basic script to take advantage of the Outlook object and extract the required information.  Users can specify the Outlook mail folder (extracts items all subfolders within this folder), the output file name, and the number of emails to extract. If you want all emails – leave this parameter blank.   ACL prompts uses for the required VB script arguments and displays the results.

Discussion
The ability to easily identify emails related to a particular period (from/to dates, times (e.g. late at night)) or to specific subjects or sender; and perform additional analysis on these items is not easy to do in Outlook.  The combination of a VB script and ACL produces a table which can easily be searched for keywords, sender name, or date and time ranges. 
Once the information is extract additional analysis can be performed in ACL using date files, keywords in the subject line, etc.  

Examples:
There is concern about a possible kickback between a contracting officer and a supplier that might have occurred three months ago over a contract for vehicles.  The application could easily identify all emails either from the supplier or the contracting officer; between the specified dates; which had the keyword “vehicles” or “trucks” or “auto”.
There is concern that an employee conversing with a known industrial terrorist and sending classified information to a person outside of the company.  The application can easily identify all sender information to determine if the employee has sent or received emails from the person in question.
The combination of ACL, visual basic and Outlook provides auditors and fraud investigators with a powerful tool to analyze email and other Outlook information. 

In ACL
•  Execute script Extract_Outlook_Info
•	Specify:
o	Outlook mailbox name
o	Outfile name (No spaces)
o	Number of emails to extract (leave blank for all emails)

The resulting file will be open and available for additional analysis.

Dave Coderre and Christian Lohyer
July 2013
