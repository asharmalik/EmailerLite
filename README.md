EmailerLite
===========
This Python script is used to send batch emails with template and multi-variable support.

Usage
===========
There is an example email campaign in the "Email Campaigns" folder. It shows the basic file structure expected. It is outlined below:

<h3>Email Campaigns\Campaign Name:</h3>
- Attachments - folder which contains any attachments to be sent
- template.txt - text file which shows the template to be sent to every person (% signifies variable)
- to_email.csv - Must be a comma delimited file and have headers "name" and "email"

Run the script and you will be given directions.

Screenshot
===========
![Screenshot](https://raw.githubusercontent.com/asharmalik/EmailerLite/master/screenshots/screenshot1.png)