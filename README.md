# Swing Scripts
## Getting Started
### Required Technologies
See tutorial [here](https://developers.google.com/apps-script/guides/typescript)
- [npm](https://nodejs.org/en/)
- [clasp](https://github.com/google/clasp)
- [vscode](https://code.visualstudio.com/)
- [git](https://git-scm.com/downloads)

```
clasp login
cd <YourRepoDir>
git clone https://github.com/dstbstr/Swing
cd Swing
code .
```

## Making Changes
There are multiple .clasp.json files.  Clasp will use the one named .clasp.json
So if switching from woodside to durandal, change `.clasp.json` to `.clasp_woodside.json`
and change `.clasp_durandal.json` to `.clasp.json`
```
clasp push
```


## Architecture
Google Scripts Apps are weird.  In order to add a submission trigger to a spreadsheet, the app must be on that sheet specifically.  You can also have a time-based script which lives on its own.  I wanted to be able to share common code (sheet names, month names, utils, etc.), and have a single git repo.

So I've created an independent script project (Swing Scripts), and on the Waiver sheet, I've added a script which includes this project as a library.  It simply calls the appropriate function from WaiverWatcher.ts

## New Year
This should largely happen after the first of the new year to avoid date complications.
* Create a folder in TES/Mondays for the existing year, and copy the current year files into the folder.
* Create a "Volunteer Schedule {year}" sheet
  * Copy an existing month from previous year
  * Create each month with the correct dates (or automate this later)
* Create a "Woodside Attendence {year}" sheet
  * Create Jan sheet manually
* In Woodside Swing/Waiver, create (or copy the existing and update) "Woodside Swing Waiver {year}"
  * Get the responder link
  * Right click in the browser and say "Create QR Code for this page"
  * Download QR code, then upload to the QR Codes folder
* Find the backing sheet "Woodside Waiver {year} Responses
  * Go to Extensions->App Scripts
  * In the Libraries, click the plus button to add the woodside automation scripts
    * The id can be found in .clasp.json
  * Create a function like `function onSubmit() { WoodsideAutomation.CopyLaatestWaiverToAttendance(); }`
  * Go to the Triggers tab
  * Add a trigger
    * Event is On form submit
    * Function is the onSubmit function from earlier
  * Fill out the waiver, and verify that the name appears in the attendence sheet
