# Client Excel Search

A Windows app for finding a client by name or COF NO in Excel workbooks. It shows the workbook, worksheet, and row, and can open the selected row in Excel.

## Run the client app

Build the client package with `.Build-ClientSearchExe.ps1`, then extract the generated ZIP and double-click `Client Excel Search.exe`. Keep the EXE and its PowerShell support files together. The EXE launches the app and checks GitHub for updates; choose **Yes** to install an available update. Updated scripts are downloaded over HTTPS and checked against the SHA-256 hashes in `version.json`.

Microsoft Excel desktop must be installed. The client chooses the folder to search. The app searches `.xls`, `.xlsx`, and `.xlsm` files in that folder and its subfolders.

## Publish an update

1. Edit `ClientSearchApp.ps1` or `ClientSearchWorker.ps1`.
2. Create a new version manifest, for example:

   ```powershell
   .\Publish-UpdateManifest.ps1 -Version 1.0.1
   ```

3. Commit the changed app script(s) and `version.json` to the repository's `main` branch. The next time the app starts, it will offer that version.

Only the app scripts and update metadata belong in this repository. Do not add client workbooks or search-index files.

