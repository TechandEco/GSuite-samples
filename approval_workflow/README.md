# Send emails from Google Sheets to approve budget requests

- Follow me on
  [Twitter](https://twitter.com/TechandEco)
- Visit this solution's step by step
  [instructions with screenshots here](https://medium.com/@TechandEco/workflow-to-collect-and-approve-budgets-using-apps-script-in-google-sheets-f713adba5d11)

## Introduction

- This is the 3rd of 5 Apps Script tutorials for users in Finance.

- In this sample I share how to use an Apps Script in a Google Sheet to
automatically create a budget submission form that you can share with end users,

- As responses arrive in the sheet, you can collaborate with other
reviewers to send approval emails in bulk, depending on whether you are
accepting,rejecting, or asking for more information about their request.

- The emails use a Google doc as a template that pulls information from the
sheet such as a user’s name, their email, the budget values they entered,
or special comments you wish to share.

## Try it

1. [Copy this Google Sheet](https://docs.google.com/spreadsheets/d/1j8aTL1mh2IUUSfRhUOqYmaNMqHxtWsgt1Vup63ewW5s/copy)

1. Visit custom menu **Send email** > Create form to generate a form with
   desired question types.

1. Modify response sheet with 3 new columns called **Action**(lists which
  Google Doc template to send to a user), **Comment** (remarks you wish to
  include in the email), **Status** (autopopulated by script with the name and
  date of the template sent). Follow instructions from
  [blog post](https://medium.com/@TechandEco/workflow-to-collect-and-approve-budgets-using-apps-script-in-google-sheets-f713adba5d11)
