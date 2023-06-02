using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace G1_RaspagemWeb
{
    class Program
    {

        static void Main(String[] args)    
        {
            Console.WriteLine("Iniciando extração em G1.com");

            int minNewsNumber = 30;

            IWebDriver driver = new ChromeDriver();

            driver.Navigate().GoToUrl("https://g1.globo.com/");

            // Accept Cookies
            IWebElement acceptCookies = driver.FindElement(By.CssSelector(".cookie-banner-lgpd_accept-button"));
            acceptCookies.Click();

            // Find all the child divs with class 'bstn-hl-wrapper'
            IList<IWebElement> wrapperElements = driver.FindElements(By.CssSelector("div.feed-post-body"));


            while (wrapperElements.Count < minNewsNumber)
            {
                // Scroll down to the bottom of the page
                IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
                jsExecutor.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");

                // Click on button to show more news, if there are more news to load
                // Check if an element with the specified locator exists
                By locator = By.CssSelector("a.load-more");
                bool elementExists = driver.FindElements(locator).Count > 0;

                if (elementExists)
                {
                    IWebElement loadMore = driver.FindElement(locator);
                    loadMore.Click();
                }

                wrapperElements = driver.FindElements(By.CssSelector("div.feed-post-body"));

            }




            // Create a DataTable
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Title");
            dataTable.Columns.Add("Summary");
            dataTable.Columns.Add("Link");

            // Loop through each 'bstn-hl-wrapper' element
            foreach (IWebElement wrapperElement in wrapperElements)
            {

                
                string titleChapeu = String.Empty;
                string summaryText = String.Empty;
                string relatedText = String.Empty;
                string link = String.Empty;

                titleChapeu = Program.GetElementText("span.feed-post-header-chapeu", wrapperElement);
                summaryText = Program.GetElementText("a.feed-post-link.gui-color-primary.gui-color-hover", wrapperElement);
                relatedText = Program.GetElementText("a.gui-color-primary.gui-color-hover.feed-post-body-title.bstn-relatedtext", wrapperElement);
                link = wrapperElement.FindElement(By.CssSelector("a.feed-post-link.gui-color-primary.gui-color-hover")).GetAttribute("href");

                // Add data row to the DataTable
                dataTable.Rows.Add(titleChapeu, summaryText, link);



                // Output the extracted values
                Console.WriteLine("Title: " + summaryText);
                Console.WriteLine("Summary: " + relatedText);
                Console.WriteLine("Link: " + link);
            }


            // Define the file path and name
            string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "G1_Noticias.xlsx");
            WriteDataTableToExcel(dataTable, excelFilePath);


            // Extract the text or attribute values from the element
            // Retrieve the page source
            string pageSource = driver.PageSource;

            // Remove script tags
            pageSource = Regex.Replace(pageSource, "<script.*?</script>", string.Empty, RegexOptions.Singleline);

            // Remove style tags
            pageSource = Regex.Replace(pageSource, "<style.*?</style>", string.Empty, RegexOptions.Singleline);

            // Remove header tags
            pageSource = Regex.Replace(pageSource, "<head.*?</head>", string.Empty, RegexOptions.Singleline);

            // Define the file path and name
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "G1.html");
            

            // Save the page source as an HTML file
            File.WriteAllText(filePath, pageSource);


            // Wait for a key press before ending the program
            // Console.WriteLine("Press any key to exit...");
            // Console.ReadKey();

            // Quit the WebDriver and release resources
            driver.Quit();
            driver.Dispose();
            driver.Quit();
            driver.Dispose();



        }

        public static string GetElementText(string selectorText, IWebElement webElement)
        {
            string elementText = string.Empty;


            // Find the element by its selector
            By selector = By.CssSelector(selectorText);

            // Output the result
            if (webElement.FindElements(selector).Count > 0)
            {
                // Get header Chapeau
                IWebElement foundElement = webElement.FindElement(selector);
                // Get the text value from the 'bstn-hl-title' element
                elementText = foundElement.Text;
            }


            return elementText;
        }

        public static void WriteDataTableToExcel(DataTable dataTable, string filePath)
        {
            // Create a new Excel document
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // Add a new workbook to the document
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Add a new worksheet to the workbook
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add a new sheet to the workbook
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                sheets.Append(sheet);

                // Write the data to the worksheet
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                foreach (DataRow row in dataTable.Rows)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        newCell.DataType = CellValues.String;
                        newCell.CellValue = new CellValue(row[col].ToString());
                        newRow.AppendChild(newCell);
                    }
                    sheetData.AppendChild(newRow);
                }
            }
        }


    }
}
