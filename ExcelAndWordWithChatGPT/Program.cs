using OfficeOpenXml;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using DocumentFormat.OpenXml;

class Program
{
    [STAThread]
    static async Task Main(string[] args)
    {
        Console.WriteLine("Select the Excel file...");
        string? excelFilePath = SelectExcelFile();

        if (string.IsNullOrEmpty(excelFilePath))
        {
            Console.WriteLine("No file selected. Exiting.");
            return;
        }

        Console.Write("Enter the column letter to read data from (e.g., E): ");
        string? columnLetter = Console.ReadLine()?.ToUpper();

        if (string.IsNullOrEmpty(columnLetter))
        {
            Console.WriteLine("Invalid column letter. Exiting.");
            return;
        }

        //context if you want to add for ai model to have
        string? context = "I am sending a review of coffee . Product name - CONTINENTAL SPECIALE Instant Coffee Granules 200gm Pouch .  Please consider the reviews accordingly and summarise it.";

        var columnData = ReadColumnDataFromExcel(excelFilePath, columnLetter, context);

        if (string.IsNullOrEmpty(columnData))
        {
            Console.WriteLine("No data found in the specified column. Exiting.");
            return;
        }

        string? summary = await GetSummaryFromOpenAPI(columnData,null);

        //change this after you have the summary 
        summary ??= "Dummy Summary";

        if (!string.IsNullOrEmpty(summary))
        {
            CreateWordDocument("summary.docx", summary);
            Console.WriteLine("Summary created successfully in summary.docx");
        }
        else
        {
            Console.WriteLine("Failed to get a summary. Exiting.");
        }
    }

    static string? SelectExcelFile()
    {
        Console.WriteLine("Give me the file path - ");
        var filePath = Console.ReadLine();
        return filePath;
    }

    static string ReadColumnDataFromExcel(string filePath, string columnLetter,string context = "")
    {
        var stringBuilder = new StringBuilder();
        if (!string.IsNullOrEmpty(context))
            stringBuilder.AppendLine(context);

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set the license context

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int column = GetColumnNumber(columnLetter);
            int row = 2; //skip first column for header

            while (!string.IsNullOrEmpty(worksheet.Cells[row, column].Text))
            {
                stringBuilder.AppendLine(worksheet.Cells[row, column].Text);
                row++;
            }
        }

        return stringBuilder.ToString();
    }

    static int GetColumnNumber(string columnLetter)
    {
        int columnNumber = 0;
        int factor = 1;

        for (int i = columnLetter.Length - 1; i >= 0; i--)
        {
            columnNumber += (columnLetter[i] - 'A' + 1) * factor;
            factor *= 26;
        }

        return columnNumber;
    }

   /// <summary>
   /// 
   /// </summary>
   /// <param name="input"></param>
   /// <param name="systemContext">Send this if you have to add some data for the AI System</param>
   /// <returns></returns>
    static async Task<string?> GetSummaryFromOpenAPI(string input,string? systemContext=null)
    {
        using (var client = new HttpClient())
        {
            string apiUrl = "https://api.openai.com/v1/chat/completions";
            string apiKey = "API KEY"; // Replace with your actual API key

            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

            var requestBody = new
            {
                model = "gpt-3.5-turbo",
                messages = new[]
                {
                new { role = "system", content = systemContext??"You are a helpful assistant." },
                new { role = "user", content = input }
            }
            };

            var response = await client.PostAsync(apiUrl, new StringContent(JsonConvert.SerializeObject(requestBody), Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                var responseContent = await response.Content.ReadAsStringAsync();
                var jsonResponse = JsonConvert.DeserializeObject<dynamic>(responseContent);
                var summary = jsonResponse?.choices[0].message.content.ToString().Trim();
                return summary;
            }
        }

        return null;
    }


    static void CreateWordDocument(string filePath, string content)
    {
        using (var wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());
            var para = body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(content));
        }
    }
}
