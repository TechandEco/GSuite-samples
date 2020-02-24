# Upload Google Sheets data as a table(s) into BigQuery

- Follow me on
  [Twitter](https://twitter.com/TechandEco)
- Visit this solution's step by step
  [instructions with screenshots here](https://medium.com/@TechandEco/apps-script-tutorial-upload-to-a-database-sheets-bigquery-2fee3724f3ca)

## Introduction

- This is the 5th of 5 Apps Script courses for Finance users.

- It’s common for financial spreadsheets that help us with pivotal operations
  to reach limitations when they begin to take the place of a database.
  Performance issues begin to be noticed such as speed or in usability such as
  search.

- This is where Google’s BigQuery database comes in, especially to help
 centralize data from multiple sources or process very large volumes of rows
 and columns when data exceeds spreadsheet quotas. It is incredibly fast, and
 organizations love to build reporting prototypes especially with Google’s
 free Data Studio dashboards because of BigQuery’s very large free monthly tier
(10GB of data).

- In this article we will review how to push data from multiple spreadsheets
  simultaneously into a database using an Apps Script.

## Try it

1. [Copy this Google Sheet](https://docs.google.com/spreadsheets/d/1MQZyW5Ds64DrtzTQhhzvtPAQtruq3EDBfaeJBG7YOe8/copy)

1. [Visit BigQuery](https://console.cloud.google.com/bigquery) and either
  login to your account or use the sandbox (if it’s your first time).

1. Enter your BigQuery project ID, dataset name, and table name in your Google
   Sheet and choose the menu Sheets to **BigQuery** > **Upload**.

  >Note: you can create new tables or add to existing ones by using the Append
  >column in the sheet.

1. For step by step instructions visit
   [blog post](https://medium.com/@TechandEco/apps-script-tutorial-upload-to-a-database-sheets-bigquery-2fee3724f3ca)
