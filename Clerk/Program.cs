using System.Diagnostics;

using Microsoft.Win32;

using Spectre.Console;

namespace Clerk;

public class Program
{
    private const string SmtpServer = "mail.kurumsaleposta.com";
    private const string Pop3Server = "mail.kurumsaleposta.com";
    private const int Pop3Port = 110;
    private const int SmtpPort = 587;

    public static int Main(string[] args)
    {
        if (!OperatingSystem.IsWindows())
        {
            AnsiConsole.MarkupLine("[red]This application only works on Windows.[/]");
            return 1;
        }

        AnsiConsole.Write(
            new FigletText("Clerk")
                .Color(Color.Cyan1));

        AnsiConsole.MarkupLine("[cyan]Outlook POP3 Profile Creator[/]");
        AnsiConsole.WriteLine();

        // Get email from user
        var email = AnsiConsole.Prompt(
            new TextPrompt<string>("Enter your [green]email address[/]:")
                .PromptStyle("yellow")
                .ValidationErrorMessage("[red]Invalid email format[/]")
                .Validate(e =>
                {
                    if (string.IsNullOrWhiteSpace(e))
                        return ValidationResult.Error("[red]Email cannot be empty[/]");

                    if (!e.Contains('@') || !e.Contains('.'))
                        return ValidationResult.Error("[red]Invalid email format[/]");

                    return ValidationResult.Success();
                }));

        // Get password from user
        var password = AnsiConsole.Prompt(
            new TextPrompt<string>("Enter your [green]password[/]:")
                .PromptStyle("yellow")
                .Secret());

        var profileName = email.Split('@')[0];
        var displayName = profileName;

        AnsiConsole.WriteLine();

        // Show configuration summary
        var table = new Table()
            .Border(TableBorder.Rounded)
            .BorderColor(Color.Cyan1)
            .AddColumn(new TableColumn("[yellow]Setting[/]").Centered())
            .AddColumn(new TableColumn("[yellow]Value[/]").Centered());

        table.AddRow("Profile Name", profileName);
        table.AddRow("Email", email);
        table.AddRow("POP3 Server", $"{Pop3Server}:{Pop3Port}");
        table.AddRow("SMTP Server", $"{SmtpServer}:{SmtpPort}");
        table.AddRow("Leave mail on server", "[red]No[/]");

        AnsiConsole.Write(table);
        AnsiConsole.WriteLine();

        // Confirm creation
        if (!AnsiConsole.Confirm("Create this Outlook profile?"))
        {
            AnsiConsole.MarkupLine("[yellow]Operation cancelled.[/]");
            return 0;
        }

        AnsiConsole.WriteLine();

        // Create the profile
        try
        {
            AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .SpinnerStyle(Style.Parse("cyan"))
                .Start("Creating Outlook profile...", ctx =>
                {
                    CreateOutlookProfileViaPRF(profileName, displayName, email, password);
                });

            AnsiConsole.WriteLine();
            AnsiConsole.Write(
                new Panel("[green]Profile created successfully![/]")
                    .Border(BoxBorder.Rounded)
                    .BorderColor(Color.Green)
                    .Header("[green]Success[/]")
                    .HeaderAlignment(Justify.Center));

            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[yellow]Next steps:[/]");
            AnsiConsole.MarkupLine("1. Open Outlook");
            AnsiConsole.MarkupLine($"2. Select the '[cyan]{profileName}[/]' profile if prompted");
            AnsiConsole.MarkupLine("3. Your email should be configured and ready");

            return 0;
        }
        catch (Exception ex)
        {
            AnsiConsole.WriteLine();
            AnsiConsole.Write(
                new Panel($"[red]{ex.Message}[/]")
                    .Border(BoxBorder.Rounded)
                    .BorderColor(Color.Red)
                    .Header("[red]Error[/]")
                    .HeaderAlignment(Justify.Center));

            return 1;
        }
    }

    private static void CreateOutlookProfileViaPRF(string profileName, string displayName, string email, string password)
    {
        // Create a PRF (Profile) file for Outlook
        var prfPath = Path.Combine(Path.GetTempPath(), $"{profileName}.prf");

        var prfContent = $@"[General]
Custom=1
ProfileName={profileName}
DefaultProfile=Yes
OverwriteProfile=Yes
ModifyDefaultProfileIfPresent=TRUE

[Service List]
ServicePOP=Microsoft POP3

[ServicePOP]
ServiceName=Microsoft POP3
AccountName={email}
EmailAddress={email}
UserName={email}
Password={password}
DisplayName={displayName}
POP3Server={Pop3Server}
POP3Port={Pop3Port}
SMTPServer={SmtpServer}
SMTPPort={SmtpPort}
SMTPAuthentication=1
SMTPUseAuth=1
SMTPUserName={email}
LeaveOnServer=0
ConnectionType=2
";

        File.WriteAllText(prfPath, prfContent);

        try
        {
            // Find Outlook executable
            var outlookPath = FindOutlookPath();
            if (string.IsNullOrEmpty(outlookPath))
            {
                throw new FileNotFoundException("Microsoft Outlook is not installed or could not be found.");
            }

            // Import PRF using Outlook
            var processInfo = new ProcessStartInfo
            {
                FileName = outlookPath,
                Arguments = $"/importprf \"{prfPath}\"",
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            using var process = Process.Start(processInfo);
            if (process != null)
            {
                // Wait for profile import to complete
                process.WaitForExit(30000);
            }

            // Give it a moment to finish writing to registry
            Thread.Sleep(2000);

            // Configure additional settings via registry
            ConfigureAccountSettings(profileName, email, password);
        }
        finally
        {
            // Clean up PRF file
            if (File.Exists(prfPath))
            {
                try { File.Delete(prfPath); } catch { }
            }
        }
    }

    private static string? FindOutlookPath()
    {
        // Check common Outlook locations
        string[] possiblePaths =
        [
            @"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
            @"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
            @"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
            @"C:\Program Files\Microsoft Office 365\root\Office16\OUTLOOK.EXE",
            @"C:\Program Files (x86)\Microsoft Office 365\root\Office16\OUTLOOK.EXE"
        ];

        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
            {
                return path;
            }
        }

        // Try to find via registry
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE");
            var path = key?.GetValue("")?.ToString();
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                return path;
            }
        }
        catch { }

        return null;
    }

    private static void ConfigureAccountSettings(string profileName, string email, string password)
    {
        var profilesPath = @"Software\Microsoft\Office\16.0\Outlook\Profiles";
        var profilePath = $@"{profilesPath}\{profileName}";

        // Find the account manager key
        using var profileKey = Registry.CurrentUser.OpenSubKey(profilePath);
        if (profileKey == null)
        {
            // Profile wasn't created by PRF import, create it manually
            CreateProfileManually(profileName, email, password);
            return;
        }

        // Look for POP3 account settings and update them
        foreach (var subKeyName in profileKey.GetSubKeyNames())
        {
            using var subKey = Registry.CurrentUser.OpenSubKey($@"{profilePath}\{subKeyName}", true);
            if (subKey == null) continue;

            var valueNames = subKey.GetValueNames();

            // Check if this is an account key
            if (valueNames.Contains("POP3 Server") || valueNames.Contains("Account Name"))
            {
                // Update leave on server settings
                subKey.SetValue("Leave Mail On Server", 0, RegistryValueKind.DWord);
                subKey.SetValue("Remove When Deleted", 1, RegistryValueKind.DWord);
                subKey.SetValue("Remove When Expired", 1, RegistryValueKind.DWord);
            }
        }
    }

    private static void CreateProfileManually(string profileName, string email, string password)
    {
        var profilesPath = @"Software\Microsoft\Office\16.0\Outlook\Profiles";
        var newProfilePath = $@"{profilesPath}\{profileName}";

        // Create profile key
        using var profilesKey = Registry.CurrentUser.CreateSubKey(profilesPath);
        if (profilesKey == null)
        {
            throw new InvalidOperationException("Failed to access Outlook profiles registry key.");
        }

        // Delete existing profile if any
        var existingSubKeys = profilesKey.GetSubKeyNames();
        if (existingSubKeys.Contains(profileName))
        {
            profilesKey.DeleteSubKeyTree(profileName);
        }

        // Create new profile
        using var newProfileKey = Registry.CurrentUser.CreateSubKey(newProfilePath);
        if (newProfileKey == null)
        {
            throw new InvalidOperationException("Failed to create profile registry key.");
        }

        // Create account manager key (9375CFF0413111d3B88A00104B2A6676)
        var accountManagerPath = $@"{newProfilePath}\9375CFF0413111d3B88A00104B2A6676";
        using var accountManagerKey = Registry.CurrentUser.CreateSubKey(accountManagerPath);
        if (accountManagerKey != null)
        {
            accountManagerKey.SetValue("NextAccountID", 2, RegistryValueKind.DWord);
            accountManagerKey.SetValue("Account Name", email, RegistryValueKind.String);

            // Create account subkey
            var accountPath = $@"{accountManagerPath}\00000001";
            using var accountKey = Registry.CurrentUser.CreateSubKey(accountPath);
            if (accountKey != null)
            {
                accountKey.SetValue("Account Name", email, RegistryValueKind.String);
                accountKey.SetValue("Display Name", email.Split('@')[0], RegistryValueKind.String);
                accountKey.SetValue("Email", email, RegistryValueKind.String);
                accountKey.SetValue("POP3 Server", Pop3Server, RegistryValueKind.String);
                accountKey.SetValue("POP3 Port", Pop3Port, RegistryValueKind.DWord);
                accountKey.SetValue("POP3 User Name", email, RegistryValueKind.String);
                accountKey.SetValue("POP3 Use SSL", 0, RegistryValueKind.DWord);
                accountKey.SetValue("SMTP Server", SmtpServer, RegistryValueKind.String);
                accountKey.SetValue("SMTP Port", SmtpPort, RegistryValueKind.DWord);
                accountKey.SetValue("SMTP User Name", email, RegistryValueKind.String);
                accountKey.SetValue("SMTP Use SSL", 0, RegistryValueKind.DWord);
                accountKey.SetValue("SMTP Use Auth", 1, RegistryValueKind.DWord);
                accountKey.SetValue("Leave Mail On Server", 0, RegistryValueKind.DWord);
                accountKey.SetValue("Remove When Deleted", 1, RegistryValueKind.DWord);
                accountKey.SetValue("Remove When Expired", 1, RegistryValueKind.DWord);
            }
        }
    }
}
