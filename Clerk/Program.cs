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
                    CreateOutlookProfile(profileName, email, password);
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
            AnsiConsole.MarkupLine("3. Enter your password when requested");

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

    private static void CreateOutlookProfile(string profileName, string email, string password)
    {
        // Check if running on Windows
        if (!OperatingSystem.IsWindows())
        {
            throw new PlatformNotSupportedException("This application only works on Windows.");
        }

        var profilesPath = @"Software\Microsoft\Office\16.0\Outlook\Profiles";
        var newProfilePath = $@"{profilesPath}\{profileName}";

        // Open or create the profiles key
        using var profilesKey = Registry.CurrentUser.CreateSubKey(profilesPath);
        if (profilesKey == null)
        {
            throw new InvalidOperationException("Failed to access Outlook profiles registry key.");
        }

        // Check if profile exists and delete it
        var existingSubKeys = profilesKey.GetSubKeyNames();
        if (existingSubKeys.Contains(profileName))
        {
            profilesKey.DeleteSubKeyTree(profileName);
        }

        // Create new profile
        using var newProfileKey = Registry.CurrentUser.CreateSubKey(newProfilePath);
        if (newProfileKey == null)
        {
            throw new InvalidOperationException("Failed to create new profile registry key.");
        }

        // Create account settings
        var accountGuid = Guid.NewGuid().ToString("B").ToUpper();
        var accountPath = $@"{newProfilePath}\{accountGuid}";

        using var accountKey = Registry.CurrentUser.CreateSubKey(accountPath);
        if (accountKey == null)
        {
            throw new InvalidOperationException("Failed to create account registry key.");
        }

        // Set account properties
        accountKey.SetValue("Account Name", email, RegistryValueKind.String);
        accountKey.SetValue("Email", email, RegistryValueKind.String);
        accountKey.SetValue("Display Name", profileName, RegistryValueKind.String);

        // POP3 Server settings
        accountKey.SetValue("POP3 Server", Pop3Server, RegistryValueKind.String);
        accountKey.SetValue("POP3 Port", Pop3Port, RegistryValueKind.DWord);
        accountKey.SetValue("POP3 Use SSL", 0, RegistryValueKind.DWord);

        // SMTP Server settings
        accountKey.SetValue("SMTP Server", SmtpServer, RegistryValueKind.String);
        accountKey.SetValue("SMTP Port", SmtpPort, RegistryValueKind.DWord);
        accountKey.SetValue("SMTP Use SSL", 0, RegistryValueKind.DWord);
        accountKey.SetValue("SMTP Use Auth", 1, RegistryValueKind.DWord);

        // User credentials
        accountKey.SetValue("POP3 User", email, RegistryValueKind.String);
        accountKey.SetValue("SMTP User", email, RegistryValueKind.String);

        // Don't leave messages on server
        accountKey.SetValue("Leave Mail On Server", 0, RegistryValueKind.DWord);
        accountKey.SetValue("Remove When Deleted", 1, RegistryValueKind.DWord);
        accountKey.SetValue("Remove When Expired", 1, RegistryValueKind.DWord);

        // Account type
        accountKey.SetValue("Account Type", "POP3", RegistryValueKind.String);
    }
}
