# Client Excel Search

A Windows app for finding a client by name or COF NO in Excel workbooks. It shows the workbook, worksheet, and row, and can open the selected row in Excel.

## Run the client app

Build with `Build-ClientSearchExe.ps1` and send the client only `Client Excel Search.exe` from the generated version folder. The EXE contains the app and support files. On first launch, it extracts them under `%LOCALAPPDATA%\ClientExcelSearch\App`, then checks GitHub for script updates. Choose **Yes** to install an available update. Updates are downloaded over HTTPS and checked against the SHA-256 hashes in `version.json`.

Requirements: 64-bit Windows, Microsoft Excel desktop, and internet access for update checks. Use **Choose file...** to pick any workbook in the folder to search; the app searches all `.xls`, `.xlsx`, and `.xlsm` files in that folder and its subfolders.

## Publish an update

1. Edit `ClientSearchApp.ps1` or `ClientSearchWorker.ps1`.
2. Create a new version manifest, for example:

   ```powershell
   .\Publish-UpdateManifest.ps1 -Version 1.0.3
   ```

3. Commit the changed app script(s) and `version.json` to the repository's `main` branch. The next time the app starts, it will offer that version.

Only the app scripts and update metadata belong in this repository. Do not add client workbooks or search-index files.

