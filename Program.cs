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

            int minNewsNumber = 20;

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
            dataTable.Columns.Add("Related Text");
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
                link = wrapperElement.FindElement(By.CssSelector("a.feed-post-link.gui-color-primary.gui-color-hover")).GetAttribute("href");

                // Find the element by its selector
                By selector = By.CssSelector("ul.bstn-relateditems");

                // Output the result
                if (wrapperElement.FindElements(selector).Count > 0)
                {
                    IList<IWebElement> relatedTexts = wrapperElement.FindElements(By.CssSelector("a.gui-color-primary.gui-color-hover.feed-post-body-title.bstn-relatedtext"));
                    foreach(IWebElement relatedTextEl in relatedTexts)
                    {
                        relatedText += relatedTextEl.Text;
                    }

                }

                // Add data row to the DataTable
                dataTable.Rows.Add(titleChapeu, summaryText, relatedText, link);

                // Output the extracted values
                Console.WriteLine("Title: " + titleChapeu);
                Console.WriteLine("Summary: " + summaryText);
                Console.WriteLine("Related Text: " + relatedText);
                Console.WriteLine("Link: " + link);
            }


            // Define the file path and name
            string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "G1_Noticias.xlsx");
            if (File.Exists(excelFilePath))
                File.Delete(excelFilePath);

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
            //Console.WriteLine("Press any key to exit...");
            //Console.ReadKey();

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
                IWebElement foundElement = webElement.FindElement(selector);
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

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Add Headers
                Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                foreach (DataColumn col in dataTable.Columns)
                {
                    Cell newCell = new Cell();
                    newCell.DataType = CellValues.String;
                    newCell.CellValue = new CellValue(col.ColumnName);
                    headerRow.AppendChild(newCell);
                }
                sheetData.AppendChild(headerRow);

                // Write the data to the worksheet
                foreach (DataRow row in dataTable.Rows)
                {
                    Row newRow = new Row();
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
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
