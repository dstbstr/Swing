# Swing Scripts

Google Apps Script automation for Woodside Swing operations.

## Requirements

- [Node.js + npm](https://nodejs.org/en/)
- [clasp](https://github.com/google/clasp)
- [VS Code](https://code.visualstudio.com/)
- [git](https://git-scm.com/downloads)

## Initial Setup

See Google's TypeScript + Apps Script guide [here](https://developers.google.com/apps-script/guides/typescript).

```bash
clasp login
git clone https://github.com/dstbstr/Swing
cd Swing
npm install
code .
```

## Build Workflow

This project builds TypeScript into a single Apps Script-compatible bundle.

- `npm run typecheck`: runs TypeScript checks only (`--noEmit`)
- `npm run build`: type-checks, bundles with esbuild, writes output to `out/`

Build artifacts:

- `out/Code.js`: bundled script for GAS
- `out/appsscript.json`: copied manifest

### Build In VS Code

The workspace has a default build task in `.vscode/tasks.json`.

- Run Build Task: `Ctrl+Shift+B`
- Task command: `npm --prefix "${workspaceFolder}" run build`

## Deploy Workflow

`.clasp.json` currently points to `rootDir: "./out"`, so push deploys built files from `out/`.

```bash
npm run build
clasp push
```

## Switching Script Targets

There are multiple clasp config files. clasp always uses the file named `.clasp.json`.

If switching targets (for example woodside vs durandal), rename files so the desired config is named `.clasp.json` before pushing.

## Callable GAS Entry Points

The bundle exposes these top-level functions for Apps Script triggers/library calls:

- `SendMonthlyReport`
- `EnsureThisMonth`
- `EnsureNextMonth`
- `UpdateVolunteers`
- `CopyLatestWaiverToAttendance`
- `CopyLaatestWaiverToAttendance` (legacy typo compatibility)

## Architecture Notes

Apps Script projects are container-bound for some trigger types (for example form submission handlers), while other automation is time-driven. This repo keeps shared logic in one codebase, then deploys as an Apps Script project/library that container-bound scripts can call.

## New Year Checklist

This should mostly happen after Jan 1 to avoid date edge cases.

- Create a folder in TES/Mondays for the existing year and copy current year files into it.
- Create a `Volunteer Schedule {year}` sheet.
- Create each month tab (or automate later).
- Create a `Woodside Attendence {year}` sheet.
- Create Jan manually.
- In Woodside Swing/Waiver, create or copy/update `Woodside Swing Waiver {year}`.
- Get the responder link.
- In the browser, select `Create QR Code for this page` and upload the image to QR Codes.
- Find the backing sheet `Woodside Waiver {year} Responses`.
- Open `Extensions -> Apps Script`.
- In Libraries, add this automation script project (script id is in `.clasp.json`).
- Add a function similar to `function onSubmit() { WoodsideAutomation.CopyLaatestWaiverToAttendance(); }`.
- In Triggers, add an `On form submit` trigger that calls `onSubmit`.
- Submit a waiver and verify the name appears in attendance.
