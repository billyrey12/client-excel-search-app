using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class ClientExcelSearchLauncher
{
    [STAThread]
    private static void Main()
    {
        string appFolder = AppDomain.CurrentDomain.BaseDirectory;
        string powerShell = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows),
            @"System32\WindowsPowerShell\v1.0\powershell.exe");

        try
        {
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

    private static void RunScript(string powerShell, string appFolder, string fileName)
    {
        string scriptPath = Path.Combine(appFolder, fileName);
        if (!File.Exists(scriptPath))
            throw new FileNotFoundException("Required file is missing: " + fileName, scriptPath);
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

