# Willow Park Checks — OneDrive Edition

Willow Park Checks is the daily premises risk-assessment app for Willow Park Montessori Nursery & Preschool.

## How records are saved

When a staff member submits an assessment, the app:

1. Saves the complete assessment locally on the tablet in IndexedDB.
2. Generates a readable PDF automatically.
3. Uploads the PDF directly to the Willow Park Microsoft OneDrive using Microsoft Graph.
4. Marks the local assessment as synced only after OneDrive confirms the upload.
5. Keeps a failed/offline upload pending so it can be retried later.


## OneDrive location

The app uses the Microsoft Graph `Files.ReadWrite.AppFolder` delegated permission. This limits the app to its own OneDrive application folder rather than granting access to the whole drive.

PDFs are filed automatically as:

`Apps/Willow Park Checks/Risk Assessments/<Academic Year>/<Term>/<Area>/<DD-MM-YYYY - Area - Staff Member.pdf>`

Example:

`Apps/Willow Park Checks/Risk Assessments/2026-2027/Autumn Term/Baby Room/04-09-2026 - Baby Room - Maura McMahon.pdf`

## Microsoft configuration

- Application name: `Willow Park Checks`
- Application (client) ID: `16e5590c-bd06-4783-8951-57235e5c6fb4`
- Redirect URI: `https://willowparkmontessori.github.io/WillowParkChecks/`
- Authentication: Single-page application / authorization code flow with PKCE through MSAL Browser
- Authority: Microsoft personal accounts (`consumers`)
- Delegated Graph permissions: `Files.ReadWrite.AppFolder`, `User.Read`
- No client secret is used or required in this public browser app.

## GitHub safety

This repository is for app code only. Do not upload completed PDFs, JSON backups, passwords, access tokens or nursery records to GitHub.

## First test after publishing

1. Open the live GitHub Pages app.
2. Open Settings → Microsoft OneDrive.
3. Tap **Connect OneDrive** and sign in with the Willow Park Microsoft account.
4. Accept the requested app-folder permission if Microsoft asks for consent.
5. Complete one clearly labelled TEST assessment.
6. Confirm the completion screen says the PDF was backed up to Willow Park OneDrive.
7. In OneDrive, verify the PDF exists under `Apps/Willow Park Checks/Risk Assessments/...` and opens correctly.

