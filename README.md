# BatchMailer
Dynamically sends emails based on a google sheet filled with data of name and email.

# Setup
- Build the project in debug mode
- Copy both the EmailTemplate.html and Contacts.xlsx into the Batchmailer/bin/Debug folder 

# QnA
## How do I insert data into BatchMailer?
- You open the file "Contacts.xlsx" in excel and input up to 500 rows of data with name and email.
- Save excel file and run BatchMailer.

## How do I change the way the email looks?
This can be changed by opening up the "EmailTemplate.html" file using html and inline css
Google chrome is the best way to view the html and see how it will look it the email (Not perfect but better than sending yourself 100 emails)
Its probably worth seeing if one of those sites like squarespace would easily make html that would look the same as developed in an email.

## How do I hookup the senders (your) email address to BatchMailer?
You will need to change the fromAddress within the code and setup 2 factor authentication. Then you will be able to create an app password (https://myaccount.google.com/apppasswords) and replace "fromPassword" in code with the new app password.

# Notes
Normal gmail accounts are limited to only sending 500 messages in a 24 hour period. A google workspace account is able to send 2000 in 24 hours.
Both the EmailTemplate.html and Contacts.xlsx need to be in Batchmailer/bin/Debug folder 