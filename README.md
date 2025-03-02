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