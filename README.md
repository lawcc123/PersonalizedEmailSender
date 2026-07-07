# WPF Front-End Demo

This folder contains a C# WPF front-page demo only. It does not run the email/PDF automation yet.

Files:

- `PersonalizedEmailSender.sln`: Visual Studio solution.
- `PersonalizedEmailSender.csproj`: WPF project file.
- `App.xaml` / `App.xaml.cs`: WPF app startup.
- `MainWindow.xaml`: WPF/XAML front page concept for a future C# desktop app.
- `MainWindow.xaml.cs`: C# code-behind with placeholder button actions.
- `preview.html`: browser-viewable mockup of the same front page.

The front page gives the user three choices:

1. Create a new personalized email sending
2. Continue the existing email work
3. View history

To run the WPF app in Visual Studio, open `PersonalizedEmailSender.sln` and press `F5`.

To run it in Visual Studio Code:

1. Install the `.NET SDK`, not only the runtime.
2. Install the `C#` or `C# Dev Kit` extension.
3. Restart VS Code after installing .NET.
4. Open the project folder.
5. Use `Run and Debug` > `Run WPF Frontend Demo`.

You can also run it from a terminal:

```powershell
dotnet run --project .\PersonalizedEmailSender\PersonalizedEmailSender.csproj
```

The project currently targets `net10.0-windows`, matching the .NET SDK installed on this machine.

Release build
dotnet publish .\PersonalizedEmailSender.csproj -c Release -r win-x64 --self-contained true -o ..\publish\PersonalizedEmailSender
