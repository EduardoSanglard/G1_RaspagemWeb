using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.IO;

namespace G1_RaspagemWeb
{
    class Program
    {

        static void Main(String[] args)    
        {
            Console.WriteLine("Iniciando extração em G1.com");

            //Asks for user input 
            Console.WriteLine("test");

            IWebDriver driver = new ChromeDriver();

            driver.Navigate().GoToUrl("https://g1.globo.com/");

            // Accept Cookies
            IWebElement acceptCookies = driver.FindElement(By.CssSelector(".cookie-banner-lgpd_accept-button"));
            acceptCookies.Click();

            // Find all the child divs with class 'bstn-hl-wrapper'
            IList<IWebElement> wrapperElements = driver.FindElements(By.CssSelector("div.feed-post-body"));

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

                // Output the extracted values
                Console.WriteLine("Title: " + summaryText);
                Console.WriteLine("Summary: " + relatedText);
                Console.WriteLine("Link: " + link);
            }


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
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "page.html");


            // Save the page source as an HTML file
            File.WriteAllText(filePath, pageSource);


            // Wait for a key press before ending the program
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();

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


    }
}
