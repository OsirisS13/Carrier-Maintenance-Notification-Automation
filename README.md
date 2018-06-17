# Carrier Maintenance Notification Automation
Python script to search an Outlook mailbox for a telecommunications carrier maintenance notification email, extract the information within and create an Outlook Calendar event for it.

In the initial implementation, circuit information is stored in a database and so once the relevant information has been extracted from the email, there is a function for connecting to and retrieving information from the database.  This could be extended to pull information from any number of other sources (Excel file, custom tool etc).  

Supported carriers so far are:

* Interoute
* euNetworks
* Level3
* Telia

Pull requests for extensions to supported carriers, sources of circuit information or any other improvements are welcomed and encouraged.

# NOTE
This script relies heavily on regex to extract the needed information, and is therefore quite fragile and sensitive to any formatting changes that are made to the notification emails. 

Non-standard modules required are:
* win32com.client (for interfacing with Outlook)
* pyodbc (for connecting to a Database)
