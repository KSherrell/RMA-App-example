# Return Merchandize Authorization Application example
## Description
The RMA Application is a Google Sheets database application that was created to keep track of the meter traffic through the shop and the work that was performed on each meter, from Receive to Job Complete. 


## Features
- Custom menus and user forms
    - fields will auto-fill with any existing customer data
    - easily update customer information by typing over the auto-fill
    - dependent drop-downs filter selections for the user
- Creates a unique identifier for each RMA job 
- Customers can request up to five different services per RMA job
- Users are assigned unique four-digit PINs
- Logging comments attached to activity cells when
    - RMA form is submitted
    - service item is complete
    - Job Closed (updated on re-open)
    - Job Complete
- Creates and emails a PDF of the RMA Job Information
    - an RMA bar code is created on the PDF for easy data entry
- Saves a Google Doc copy of the RMA Job Info
- Email notifications for 
    - creating an RMA
    - updating an RMA
    - generating a 'Job Closed' result from the 'Receive' dialogue
    - re-opening an RMA
    - completing a line item ('Job Complete')
    - a visitor to the Customer Login screen submits the 'I've had difficulty logging in' form 

## Screenshots
View in screenshots folder 

## Project Tech
- Javascript
- jQuery
- Google App Script (and its native editor)
- CLASP
- Visual Studio Code
- HTML 
- CSS
- Materialize.css
- Github
- Git, Windows command line
- Google Drive, Sheets, Gmail
- markdown cheatsheets&nbsp;&nbsp;&nbsp;&nbsp;: &nbsp;) 

## Working on next ...
- taking the customer data out of the RMA App and pulling it from Smith: Customer DB instead
- code optimization

## How to use this code

This code creates a Google Sheets application and utilizes App Script.

- I removed my workbook Id's and company email addresses - you will need to replace those with your own workbook Id's and email addresses
- copy the code in the SRC folder to your Sheets project
    - the .js files are the .GS files, all others are html files 
- deploy the project as a web app

The files in create-new-rma are the app script files for the internal and customer facing web forms. Copy the code to a new App Script project, personalize your variables, and deploy.


