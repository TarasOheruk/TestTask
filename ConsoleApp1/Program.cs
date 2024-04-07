using HtmlAgilityPack;
using OfficeOpenXml;

public class Program
{

    static async Task DownloadExcelFile(string url, string fileName)
    {
        using (HttpClient httpClient = new HttpClient())
        {

            try
            {
                HttpResponseMessage response = await httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();

                string excelFileUrl = await GetExcelFileUrl(responseBody);
                if (string.IsNullOrEmpty(excelFileUrl))
                {
                    Console.WriteLine("Link to Excel file not found.");
                    return;
                }

                string fullUrl = "https://bakerhughesrigcount.gcs-web.com/" + excelFileUrl;

                if (!Uri.TryCreate(fullUrl, UriKind.Absolute, out Uri absoluteUri))
                {
                    Console.WriteLine("Invalid Excel file URL.");
                    return;
                }

                HttpResponseMessage excelResponse = await httpClient.GetAsync(absoluteUri);
                excelResponse.EnsureSuccessStatusCode();

                byte[] fileBytes = await excelResponse.Content.ReadAsByteArrayAsync();

                File.WriteAllBytes(fileName, fileBytes);

                Console.WriteLine($"Excel file downloaded successfully: {fileName}");
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine($"Http request error: {e}");
            }
        }
    }

    public static async Task<string> GetExcelFileUrl(string html)
    {
        HtmlDocument htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(html);

        HtmlNode linkNode = htmlDoc.DocumentNode.SelectSingleNode("//a[contains(@title, 'Worldwide Rig Count Jan 2007_Mar 2024.xlsx')]");
        return linkNode?.GetAttributeValue("href", "");
    }

    public static async Task ConvertExcelToCSV(string fileName, string outputFileName)
    {
        try
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    await WriteDataToCSV(worksheet, outputFileName);
                }
                else
                {
                    Console.WriteLine("Sheet not found in Excel-file.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error Excel to CSV: {ex.Message}");
        }
    }

    public static async Task WriteDataToCSV(ExcelWorksheet worksheet, string outputFileName)
    {
        int currentYear = DateTime.Now.Year;
        int twoYearsAgo = currentYear - 2;
        try
        {
            int rowCount = worksheet.Dimension.Rows;
            int columnCount = worksheet.Dimension.Columns;

            using (StreamWriter writer = new StreamWriter(outputFileName))
            {
                bool startWriting = false; 

                for (int row = 1; row <= rowCount; row++)
                {
                    string date = worksheet.Cells[row, 2]?.GetValue<string>();

                    if (date != null && (date.StartsWith(currentYear.ToString())))
                    {
                        startWriting = true; 
                    }
                    else if (date != null && (date.StartsWith(twoYearsAgo.ToString())))
                        startWriting = false;

                    if (startWriting)
                    {
                        string rowData = await GetRowData(worksheet, row, columnCount);

                        await writer.WriteLineAsync(rowData);
                    }
                }
            }

            Console.WriteLine($"Succes getting data from {currentYear} and {twoYearsAgo} in file {outputFileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting data from {currentYear} and {twoYearsAgo} : {ex.Message}");
        }
        
    }

    public static async Task<string> GetRowData(ExcelWorksheet worksheet, int row, int columnCount)
    {
        string rowData = "";

        for (int col = 1; col <= columnCount; col++)
        {
            object value = worksheet.Cells[row, col].Value;

            if (value != null)
            {
                rowData += value.ToString();
            }

            if (col < columnCount-1)
            {
                rowData += ",";
            }
        }

        return rowData;
    }
    static async Task Main(string[] args)
    {
        string url = "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl";
        string fileName = "Worldwide_Rig_Count_Jan_2007_Mar_2024.xlsx";
        string outputFileName = "output.csv";

        await DownloadExcelFile(url, fileName);
        await ConvertExcelToCSV(fileName, outputFileName);
    }
}
