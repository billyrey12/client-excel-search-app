# Client Excel Search

A Windows app for finding a client by name or COF NO in Excel workbooks. It shows the workbook, worksheet, and row, and can open the selected row in Excel.

## Start the app

Double-click `Start Client Search.cmd`. The launcher checks GitHub for a newer version before opening the app. When an update is available, choose **Yes** to install it. Updates are downloaded over HTTPS and checked against the SHA-256 hashes in `version.json`.

Microsoft Excel must be installed. The app searches `.xls`, `.xlsx`, and `.xlsm` files in the selected folder and its subfolders.

## Publish an update

1. Edit `ClientSearchApp.ps1` or `ClientSearchWorker.ps1`.
2. Create a new version manifest, for example:

   ```powershell
   .\Publish-UpdateManifest.ps1 -Version 1.0.1
   ```

3. Commit the changed app script(s) and `version.json` to the repository's `main` branch. The next time the app starts, it will offer that version.

Only the app scripts and update metadata belong in this repository. Do not add client workbooks or search-index files.

