# AutoHomeworkSolver
A C# application that automates homework solving by downloading assignments from a platform, sending them to Google Gemini API for solutions, saving the results as Word/PDF files, and uploading them back automatically.
## 📦 Requirements

Before running the project, ensure that you have the following installed:

- [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) (Required to run the project)
- A Google Cloud account with the Gemini API enabled to obtain your API key
- Visual Studio or any editor that supports C# and .NET

## ⚙️ Setup

1. Clone the repository:

```bash
git clone https://github.com/HussamAlmaaitah/AutoHomeworkSolver.git
```

2. Open the project in Visual Studio or your preferred editor.

3. Replace the `GeminiApiKey` constant with your own API key in the code:

```csharp
private const string GeminiApiKey = "Enter_Your_API_Key_Here";
```

4. Build and run the project.

## 🛠️ Features

- Automatically logs into the homework platform.
- Downloads assignment files.
- Uses the Gemini API to solve the assignments.
- Saves the answers in Word/PDF files.
- Re-uploads the files automatically.
