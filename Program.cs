using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using Aspose.Words;
using DocumentFormat.OpenXml;
using OpenQA.Selenium.Support.UI;
using Microsoft.Win32;


namespace AutoHomeworkSolver
{
    internal class Program
    {
        private const string GeminiApiKey = "AIzaSyAQI8zX06gC6gQx8IsmwE5B8vQFjPFkrGg";

        static async Task Main(string[] args)
        {
            Console.SetError(TextWriter.Null);

            // Welcome message
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("╔════════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║                AUTOMATIC HOMEWORK SOLVER                       ║");
            Console.WriteLine("╚════════════════════════════════════════════════════════════════╝");
            Console.ResetColor();

            // Get login credentials
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write(" Enter your student number: ");
            Console.ResetColor();
            string studentNumber = Console.ReadLine().Trim();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write(" Enter your password: ");
            Console.ResetColor();
            string password = Console.ReadLine().Trim();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write(" Enter assignment URL: ");
            Console.ResetColor();
            string urlAssignment = Console.ReadLine().Trim();
            string projectDirectory = Directory.GetParent(AppContext.BaseDirectory).Parent.Parent.Parent.FullName;
            string downloadPath = Path.Combine(projectDirectory, "Downloads");

            Directory.CreateDirectory(downloadPath);

            // Starting browser
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\n◉ Initializing browser...");
            Console.ResetColor();

            var options = new ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", downloadPath);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddUserProfilePreference("plugins.always_open_pdf_externally", true);
            options.AddArgument("--log-level=1");
            options.AddExcludedArgument("enable-logging");
            Environment.SetEnvironmentVariable("GLOG_minloglevel", "2");
            options.AddArgument("--disable-blink-features=AutomationControlled");
            options.AddArgument("--disable-features=SafeBrowsing,PasswordLeakToggleMove,PasswordLeakDetection");
            var prefs = new Dictionary<string, object>
            {
                // Disable Chrome’s offer-to-save/password-manager bubble
                ["credentials_enable_service"] = false,
                ["profile.password_manager_enabled"] = false,

                // Disable Safe Browsing-based leak warnings
                ["safebrowsing.enabled"] = false
            };
            options.AddUserProfilePreference("prefs", prefs);
            options.AddArgument("--disable-infobars");
            options.AddArgument("--disable-extensions");
            options.AddArgument("--no-sandbox");
            options.AddArgument("--disable-dev-shm-usage");
            options.AddArgument("--disable-gpu");
            options.AddArgument("--disable-software-rasterizer");
            options.AddArgument("--disable-notifications");
            options.AddExcludedArgument("enable-automation");
            options.AddAdditionalChromeOption("useAutomationExtension", false);
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddUserProfilePreference("safebrowsing.enabled", false);


            options.PageLoadStrategy = PageLoadStrategy.Normal;
            options.UnhandledPromptBehavior = UnhandledPromptBehavior.Ignore;

            IWebDriver driver = new ChromeDriver(options);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ Navigating to assignment URL...");
            Console.ResetColor();

            driver.Navigate().GoToUrl(urlAssignment);
            IWebElement snum = driver.FindElement(By.Id("username"));
            snum.SendKeys(studentNumber);
            IWebElement spass = driver.FindElement(By.Id("password"));
            spass.SendKeys(password);
            IWebElement enter = driver.FindElement(By.Id("loginbtn"));
            enter.Submit();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ Successfully logged in");
            Console.WriteLine("◉ Locating assignment file...");
            Console.ResetColor();

            IWebElement block = driver.FindElement(By.ClassName("fileuploadsubmission"));
            IWebElement fileAssignment = block.FindElement(By.TagName("a"));
            string fileName = fileAssignment.Text.Trim();
            fileAssignment.Click();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ Downloading assignment file: " + fileName);
            Console.ResetColor();

            string downloadedFilePath = WaitForDownloadedFile(downloadPath, fileName);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ File downloaded successfully");
            Console.WriteLine("◉ File saved at: " + downloadedFilePath);
            Console.ResetColor();

            // Prompt selection
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("╔════════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║                       PROMPT SELECTION                         ║");
            Console.WriteLine("╚════════════════════════════════════════════════════════════════╝");
            Console.ResetColor();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("1. Use default prompt ");
            Console.WriteLine("2. Custom prompt");
            Console.Write("\nYour choice (1/2): ");
            Console.ResetColor();

            string question = "Please read the attached assignment file, identify the questions, and provide only the final complete answers in a clean and organized format. No explanations needed. The answers should be ready to be submitted to the instructor as-is ";
            string c = Console.ReadLine();

            if (c != "1")
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write("Enter your custom prompt for Gemini: ");
                Console.ResetColor();
                question = Console.ReadLine();
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine("Using custom prompt");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine("Using default prompt");
                Console.ResetColor();
            }

            // Processing with AI
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("╔════════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║                    PROCESSING ASSIGNMENT                       ║");
            Console.WriteLine("╚════════════════════════════════════════════════════════════════╝");
            Console.ResetColor();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ Sending request to Gemini AI...");
            Console.ResetColor();

            string response = await AskGemini(question, downloadedFilePath);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ Response received successfully!");
            Console.ResetColor();

            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("╔════════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║                        AI RESPONSE                             ║");
            Console.WriteLine("╚════════════════════════════════════════════════════════════════╝");
            Console.ResetColor();

            Console.WriteLine(response);

            // Save options
            string projectDirectory2 = Directory.GetParent(AppContext.BaseDirectory).Parent.Parent.Parent.FullName;
            string solutionFolder = Path.Combine(projectDirectory2, "the solution");
            Directory.CreateDirectory(solutionFolder);
            string solutionFilePath = null;

            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("╔════════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║                    SAVE SOLUTION FORMAT                        ║");
            Console.WriteLine("╚════════════════════════════════════════════════════════════════╝");
            Console.ResetColor();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("1. Microsoft Word (.docx)");
            Console.WriteLine("2. PDF Document (.pdf)");
            Console.Write("\nYour choice (1/2): ");
            Console.ResetColor();

            string choice = Console.ReadLine().Trim();

            string baseName = Path.GetFileNameWithoutExtension(downloadedFilePath);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ Creating solution document...");
            Console.ResetColor();

            if (choice == "1")
            {
                string wordPath = Path.Combine(solutionFolder, baseName + "_solution.docx");
                CreateWord(wordPath, response);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("◉ Word document created successfully");
                Console.WriteLine("◉ Saved at: " + wordPath);
                Console.ResetColor();
                solutionFilePath = wordPath;
            }
            else if (choice == "2")
            {
                string pdfPath = Path.Combine(solutionFolder, baseName + "_solution.pdf");
                CreatePdf(pdfPath, response);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("◉ PDF document created successfully");
                Console.WriteLine("◉ Saved at: " + pdfPath);
                Console.ResetColor();
                solutionFilePath = pdfPath;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("⨯ Invalid choice, skipping file save.");
                Console.ResetColor();
            }

            // Submit solution
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("╔════════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║                    SUBMITTING SOLUTION                         ║");
            Console.WriteLine("╚════════════════════════════════════════════════════════════════╝");
            Console.ResetColor();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ Navigating to submission page...");
            Console.ResetColor();

            IWebElement AddSubmission = driver.FindElement(By.CssSelector(".btn.btn-primary"));
            AddSubmission.Click();

            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            var addButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector("a[title='Add...']")));

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ Adding solution file...");
            Console.ResetColor();

            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", addButton);

            Thread.Sleep(2000);

            var waits = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(driver => driver.FindElement(By.CssSelector("input[type='file']")).Displayed);

            IWebElement fileInput = driver.FindElement(By.CssSelector("input[type='file']"));
            fileInput.SendKeys(solutionFilePath);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("◉ Uploading solution...");
            Console.ResetColor();

            IWebElement submitBtn = driver.FindElement(By.CssSelector(".fp-upload-btn.btn-primary.btn"));
            submitBtn.Click();

            //WebDriverWait wait2 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            //IWebElement submitButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id("id_submitbutton")));

            //submitButton.Click();

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\n╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║            SOLUTION SUCCESSFULLY SUBMITTED                    ║");
            Console.WriteLine("╚═══════════════════════════════════════════════════════════════╝");
            Console.ResetColor();
        }

        private static async Task<string> AskGemini(string question, string filePath)
        {
            // إذا كان الملف وورد، حوّله أولاً إلى PDF
            string ext = Path.GetExtension(filePath).ToLower();
            if (ext == ".docx" || ext == ".doc")
            {
                filePath = ConvertWordToPdf(filePath);
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine("Converted Word to PDF: " + filePath);
                Console.ResetColor();
            }

            using var httpClient = new HttpClient();
            var parts = new object[]
            {
                new { text = question },
                new {
                    inline_data = new {
                        mime_type = "application/pdf",
                        data = Convert.ToBase64String(File.ReadAllBytes(filePath))
                    }
                }
            };
            var requestBody = new { contents = new[] { new { parts } } };

            var response = await httpClient.PostAsync(
                $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GeminiApiKey}",
                new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json")
            );

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);

            if (doc.RootElement.TryGetProperty("error", out var err))
                return $"API Error: {err.GetProperty("message").GetString()}";

            return doc.RootElement
                      .GetProperty("candidates")[0]
                      .GetProperty("content")
                      .GetProperty("parts")[0]
                      .GetProperty("text")
                      .GetString() ?? "Error: Empty response";
        }


        static string WaitForDownloadedFile(string downloadPath, string expectedFileName, int timeoutSeconds = 30)
        {
            string filePath = Path.Combine(downloadPath, expectedFileName);
            int waited = 0;

            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.Write("Waiting for download");
            Console.ResetColor();

            while (!File.Exists(filePath))
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.Write(".");
                Console.ResetColor();
                System.Threading.Thread.Sleep(1000);
                waited++;
                if (waited >= timeoutSeconds)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("\n⨯ Download timeout. Please check your connection and try again.");
                    Console.ResetColor();
                    throw new Exception("Timeout waiting for the file to be downloaded.");
                }
            }

            Console.WriteLine();
            return filePath;
        }

        static void CreateWord(string filepath, string content)
        {
            using var wordDoc = WordprocessingDocument.Create(
                filepath,
                WordprocessingDocumentType.Document
            );
            var mainPart = wordDoc.AddMainDocumentPart();
            var body = new DocumentFormat.OpenXml.Wordprocessing.Body();
            foreach (var line in content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
            {
                var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new Text(line)));
                body.Append(paragraph);
            }
            mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(body);
            mainPart.Document.Save();
        }

        static void CreatePdf(string filepath, string content)
        {
            var document = new PdfDocument();
            document.Info.Title = "Solution Document";

            var font = new XFont("Verdana", 12, XFontStyleEx.Regular);
            double margin = 40;
            double lineHeight = font.GetHeight();

            var lines = content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            PdfPage page = document.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            var tf = new XTextFormatter(gfx);
            tf.Alignment = XParagraphAlignment.Left;

            double y = margin;
            double usableHeight = page.Height - 2 * margin;

            foreach (var line in lines)
            {
                if (y + lineHeight > page.Height - margin)
                {
                    page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    tf = new XTextFormatter(gfx);
                    tf.Alignment = XParagraphAlignment.Left;
                    y = margin;
                }

                var rect = new XRect(margin, y, page.Width - 2 * margin, lineHeight);
                tf.DrawString(line, font, XBrushes.Black, rect, XStringFormats.TopLeft);
                y += lineHeight;
            }

            document.Save(filepath);
        }

        private static string GetMimeType(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLower();
            return extension switch
            {
                ".pdf" => "application/pdf",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".doc" => "application/msword",
                ".txt" => "text/plain",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                _ => "application/octet-stream"
            };
        }

        private static string ConvertWordToPdf(string wordPath)
        {
            string pdfPath = Path.Combine(
                Path.GetDirectoryName(wordPath)!,
                Path.GetFileNameWithoutExtension(wordPath) + ".pdf"
            );
            var doc = new Aspose.Words.Document(wordPath);
            doc.Save(pdfPath, Aspose.Words.SaveFormat.Pdf);
            return pdfPath;
        }
    }
}