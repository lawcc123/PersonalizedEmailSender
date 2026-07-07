# Personalized Email Sender

Personalized Email Sender is a Windows WPF desktop app for preparing and sending personalized emails through Microsoft Outlook.

The app helps the user:

1. Create a new personalized email sending job.
2. Load recipients from an Excel or CSV file.
3. Select a recipient email column.
4. Write a personalized email subject and body using merge fields such as `{{OwnerName}}`.
5. Optionally attach a Word template and generate personalized Word/PDF attachments.
6. Add common attachment files.
7. Preview generated emails before sending.
8. Send all prepared emails through Outlook.
9. Save drafts and view send history.

## Requirements

- Windows
- .NET SDK 10
- Microsoft Outlook, for opening or sending emails
- Microsoft Word, for Word template generation and PDF conversion

## Project Structure

```text
Property Management Project
+-- PersonalizedEmailSender
|   +-- PersonalizedEmailSender.csproj
|   +-- PersonalizedEmailSender.sln
|   +-- App.xaml
|   +-- MainWindow.xaml
|   +-- app source files
```

## Main Features

- New job setup window
- Recipient file loading from `.xlsx`, `.xls`, and `.csv`
- Excel worksheet selection when a workbook has multiple sheets
- Recipient preview window with the selected sending-email column highlighted
- Email subject/body merge fields
- Word template merge fields using `{{ColumnName}}`
- Optional Word-to-PDF conversion
- Attachment size validation with a 20 MB limit
- Built-in email preview window
- Send to All with progress display
- Draft saving and reopening
- Send history view
- App-managed signature text and optional signature picture

## Run The App

From the repository root:

```powershell
dotnet run --project .\PersonalizedEmailSender\PersonalizedEmailSender.csproj
```

Or open `PersonalizedEmailSender.sln` in Visual Studio and press `F5`.

## Publish Release Build

From the `PersonalizedEmailSender` folder:

```powershell
dotnet publish .\PersonalizedEmailSender.csproj -c Release -r win-x64 --self-contained true -o ..\publish\PersonalizedEmailSender
```

Copy the full published folder to another computer:

```text
publish\PersonalizedEmailSender
```

Run:

```text
PersonalizedEmailSender.exe
```

Copy the whole folder, not only the `.exe`.

## Notes About Saved Paths

The app can be copied to another computer, but saved drafts/settings may contain file paths from the original computer.

For example:

```text
C:\Users\YourName\Downloads\template.docx
```

If that file does not exist on another computer, the app should show a warning or validation error. For best results, each user should create new jobs and select files from their own computer.

## Git Notes

Build output and private working files should not be committed.

The project includes a `.gitignore` for:

- `bin/`
- `obj/`
- `publish/`
- generated binaries
- drafts/history/settings
- private Excel, CSV, Word, and PDF files
