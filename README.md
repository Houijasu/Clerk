# Clerk

A Windows command-line tool for creating Outlook POP3 email profiles.

## Features

- Interactive CLI with colored output using Spectre.Console
- Secure password input (hidden)
- Input validation for email addresses
- Configuration summary before profile creation
- Automatic registry configuration for Outlook

## Requirements

- Windows OS
- .NET 10.0 or later
- Microsoft Outlook (Office 16.0)

## Installation

```bash
git clone https://github.com/Houijasu/Clerk.git
cd Clerk
dotnet build
```

## Usage

```bash
dotnet run --project Clerk
```

The application will prompt you for:
1. Email address
2. Password

## Configuration

The tool creates Outlook profiles with the following settings:

| Setting | Value |
|---------|-------|
| POP3 Server | mail.kurumsaleposta.com |
| POP3 Port | 110 |
| SMTP Server | mail.kurumsaleposta.com |
| SMTP Port | 587 |
| Leave mail on server | No |

## Dependencies

- [Spectre.Console](https://spectreconsole.net/) - Beautiful console output
- [Spectre.Console.Cli](https://spectreconsole.net/cli/) - Command-line parsing
- [Spectre.Console.Analyzer](https://www.nuget.org/packages/Spectre.Console.Analyzer) - Code analysis

## License

MIT
