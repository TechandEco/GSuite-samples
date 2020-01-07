# Auto-calculate rooms needed for a conference 🎟️ in a Google Sheet via an Apps Script

This is a sample to show a basic custom function using an Apps Script in
Google Sheets. A
[custom function](https://developers.google.com/apps-script/guides/sheets/functions)
is a formula you create and can call into your sheet in order to perform
calculations that can otherwise be extremely manual, error prone, or take
precious time.

You can find the use case and steps in this
[blog post](https://medium.com/@TechandEco/auto-calculate-rooms-needed-for-a-conference-in-a-google-sheet-via-an-apps-script-31bcca925c34)
![Screenshot of custom function called `ROOM_TYPE`](room-type.png)

## Steps to run code

1. Create a new Google Sheet
1. Click Tools > Script editor in the Sheet's menu bar.
1. Paste code from the Code.js file into the sheet's script editor and save.
1. In the Google Sheet, if you have a the number of participants in cell C2 for
   example, you can enter =ROOM_TYPE(C2) in another cell to use the custom
   function in the script.
