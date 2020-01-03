
# Auto-calculate rooms needed for a conference 🎟️ in a Google Sheet via an Apps Script

This is a sample to show a simple custom function using an Apps Script in
Google Sheets.

You can find the use case and steps in this
[blog post](https://medium.com/@TechandEco/auto-calculate-rooms-needed-for-a-conference-in-a-google-sheet-via-an-apps-script-31bcca925c34)

![Screenshot of custom function called `ROOM_TYPE`](room-type.png)

## Steps to run code

1. Create a new Google Sheet
1. Click Tools > Script editor in the Sheet's menu bar.
1. Paste code from the Code.js file into the sheet's script editor and save.
1. Return to Google Sheet and enter `=ROOM_TYPE`(cell position of registered participants)
   in order to have the function perform it's calculation.
