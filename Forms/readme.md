# __Form__ Setup Guide

Follow this setup guide to ensure the correct fields for
cancellation request form are process correctly by the
automated cancellation request script.

1. Create all of the form fields required for this script: ![formSetup1](/Resources/formSetup1.png)
	- For the `Courses` field you need to enter all of the courses for current semester. Add and remove course as needed.
	- For the `Room Number` field you need to enter all of the room numbers for the current semester. Add and remove rooms as needed.
	- For the `Do you have more than one session question?` question setup the form to submit the form if the user enters `no`, else if they enter `yes` than go to the next section.

2. Enforce these settings for the form: ![formSetup2](/Resources/formSetup2.png)

3. `Optional:` Enable a Status Bar ![formSetup3](/Resources/formSetup3.png)

4. Create a new response spreadsheet. ![formSetup4](/Resources/formSetup4.png) ![formSetup4](/Resources/formSetup5.png)

5. Go on to [script setup guide](/Scripts/readme.md) if you have not setup the script yet.
