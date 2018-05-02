# Script Setup Guide

1. Either obtain a copy of Session Cancellation Request Form from
SI supervisor or create your own cancellation form with this [guide](/Forms/readme.md).

2. Open responses spreadsheet attached to this form. And start the script editor.
![scriptSetup1](/Resources/scriptSetup1.png)  

3. Once your are in the script editor copy the code from [onFormSubmit.gs](/Scripts/onFormSubmit.gs) and paste it into code editor.
![scriptSetup2](/Resources/scriptSetup2.png)  
	- Then enter the values for the global variables.
	```JavaScript
	var officePhoneNumber = ''; // NE SI Office Phone Number
	var supervisorName = ''; 	// Name of current SI supervisor
	var supervisorEmail = '';   // The person recieving the cancellation request. Only use a TCCD email.
	var responesFolderID = '';	// Where the responses will be stored.
	```

4. Name the project, and rename the script.
![scriptSetup3](/Resources/scriptSetup3.png)
![scriptSetup4](/Resources/scriptSetup4.png)  

5. Click edit and then current triggers.
![scriptSetup5](/Resources/scriptSetup5.png)

6. Click the link to add a trigger.
![scriptSetup6](/Resources/scriptSetup6.png)

7. Create this trigger.
![scriptSetup7](/Resources/scriptSetup7.png)  

8. Review the permissions and allow them.
![scriptSetup8](/Resources/scriptSetup8.png)  
![scriptSetup9](/Resources/scriptSetup9.png)  
