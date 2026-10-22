using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

internal static class ClientExcelSearchLauncher
{
    private static readonly string[] PayloadFiles =
    {
        "ClientSearchApp.ps1",
        "ClientSearchWorker.ps1",
        "UpdateClientSearch.ps1",
        "version.json"
    };

    [STAThread]
    private static void Main()
    {
        string appFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ClientExcelSearch",
            "App");
        string powerShell = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows),
            @"System32\WindowsPowerShell\v1.0\powershell.exe");

        try
        {
            ExtractMissingPayload(appFolder);
            RunScript(powerShell, appFolder, "UpdateClientSearch.ps1");
            RunScript(powerShell, appFolder, "ClientSearchApp.ps1");
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Client Excel Search could not start.\r\n\r\n" + ex.Message,
                "Client Excel Search",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static void ExtractMissingPayload(string appFolder)
    {
        Directory.CreateDirectory(appFolder);
        Assembly assembly = Assembly.GetExecutingAssembly();

        foreach (string fileName in PayloadFiles)
        {
            string filePath = Path.Combine(appFolder, fileName);
            if (File.Exists(filePath))
                continue;

            string resourceName = "Payload." + fileName;
            using (Stream resource = assembly.GetManifestResourceStream(resourceName))
            {
                if (resource == null)
                    throw new FileNotFoundException("Embedded app file is missing: " + fileName);
                using (FileStream output = File.Create(filePath))
                    resource.CopyTo(output);
            }
        }
    }

    private static void RunScript(string powerShell, string appFolder, string fileName)
    {
        string scriptPath = Path.Combine(appFolder, fileName);
        if (!File.Exists(scriptPath))
            throw new FileNotFoundException("Required app file is missing: " + fileName, scriptPath);
        if (!File.Exists(powerShell))
            throw new FileNotFoundException("Windows PowerShell 5.1 was not found.", powerShell);

        var startInfo = new ProcessStartInfo
        {
            FileName = powerShell,
            Arguments = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File \"" + scriptPath + "\"",
            WorkingDirectory = appFolder,
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        using (Process process = Process.Start(startInfo))
        {
            if (process == null)
                throw new InvalidOperationException("Windows PowerShell could not be started.");
            process.WaitForExit();
            if (process.ExitCode != 0)
                throw new InvalidOperationException(fileName + " exited with code " + process.ExitCode + ".");
        }
    }
}

