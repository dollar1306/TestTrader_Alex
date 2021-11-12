using System;
using System.IO;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using excel = Microsoft.Office.Interop.Excel;

/*
Maybe the code is not 100% clean and according to what you wanted, I agree.
I tried to do my best, while I had free time "at night".
 */

namespace TestTrader_Alex
{
    class Program
    {
        static IWebDriver driver;

        static void Main(string[] args)
        {
            try
            {
                driver = new ChromeDriver
                {
                    Url = GetData(1)
                };
                driver.Manage().Window.Maximize();
                LogIn();
            }
            catch (Exception ex)
            {

                Console.WriteLine("Error: " + ex.Message);
            }

        }
        //get data from excel
        public static string GetData(int index)
        {
            try
            {
                excel.Application x1app = new excel.Application();
                excel.Workbook x1workbook = x1app.Workbooks.Open(@"C:\Users\USER\source\repos\AlexTest_Trader_I_Forex\TestData.xlsx");
                excel._Worksheet x1worksheet = (excel.Worksheet)x1workbook.Sheets[1];
                excel.Range x1range = x1worksheet.UsedRange;

                string url, user, pass;
                for (int i = 2; ;)
                {
                    for (int j = index; j <= index; j++)
                    {
                        if (j == 1)
                        {
                            url = x1worksheet.Cells[i, j].Value.ToString();
                            return url;
                        }
                        else if (j == 2)
                        {
                            user = x1worksheet.Cells[i, j].Value.ToString();
                            return user;
                        }
                        else if (j == 3)
                        {
                            pass = x1worksheet.Cells[i, j].Value.ToString();
                            return pass;
                        }
                    }
                }
                x1workbook.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(value: "Error " + ex.Message);
                return null;
            }
        }

        public static void LogIn()
        {
            try
            {
                string user = GetData(2);
                IWebElement email = driver.FindElement(By.CssSelector("input[name='UserName']"));
                email.SendKeys(user);
                string pass = GetData(3);
                IWebElement password = driver.FindElement(By.CssSelector("input[name='Password']"));
                password.SendKeys(pass);
                IWebElement pushBtn = driver.FindElement(By.Id("btnOkLogin"));
                pushBtn.Click();
                driver.Navigate().Back();
                LogOut();

            }
            catch (NoSuchElementException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        //function to find and catch text
        public static void LogOut()
        {
            try
            {
                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                //Search for a question mark selector and grab text
                IWebElement element = driver.FindElement(By.CssSelector("i[class='ask ico-wb-help']"));
                string text = element.GetAttribute("title");
                Console.WriteLine("This is a Text: " + text);
                //Function call, file write
                WriteToTextFile(text);
                //call function to check loaded page

                if (WaitAndCheck())
                {
                    Console.WriteLine("The page loaded after 10 seconds");
                }
                else
                {
                    Console.WriteLine("Error, the page not loaded");
                }
            }
            catch (NoSuchElementException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

        }
        //function to wait and check page if loaded
        public static bool WaitAndCheck()
        {
            Task.Delay(10000);
            try
            {
                IWebElement check = driver.FindElement(By.CssSelector("i[class='ask ico-wb-help']"));
                bool isDisplayed = check.Displayed;
                return isDisplayed;
            }
            catch (NoSuchElementException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                return false;
            }
        }

        //function to write cath text to file
        public static void WriteToTextFile(string text)
        {
            try
            {
                //Pass the filepath and filename to the StreamWriter Constructor
                StreamWriter sw = new StreamWriter(@"C:\Users\USER\source\repos\AlexTest_Trader_I_Forex\TextFile1.txt");
                //Write a line of text
                sw.WriteLine(text);
                //Close the file
                sw.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }
    }
}
