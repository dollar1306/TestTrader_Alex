using System;
using System.Diagnostics;
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
                //WaitForLoad();//call function to wait 10 seconds
                LogIn();//call function to login
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
                IWebElement elementBtn = null;
                IWebElement pushBtn = WaitForPageLoad(elementBtn);//to test function use "btnOkLogin1"
                string user = GetData(2);
                IWebElement email = driver.FindElement(By.CssSelector("input[name='UserName']"));
                email.SendKeys(user);
                string pass = GetData(3);
                IWebElement password = driver.FindElement(By.CssSelector("input[name='Password']"));
                password.SendKeys(pass);
            
                try
                {
                    pushBtn.Click();
                }catch(NullReferenceException ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
                WaitForLoad(); // call function to wait 10 seconds
                LogOut();// call function to logout

            }
            catch (NoSuchElementException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        //function to logout and find and catch text
        public static void LogOut()
        {
            try
            {
                driver.Navigate().Back();
                IWebElement outBtn = driver.FindElement(By.CssSelector("#ExitAlertbutton0"));
                outBtn.Click();
                
                //Search for a question mark selector and grab text
                IWebElement element = driver.FindElement(By.CssSelector("i[class='ask ico-wb-help']"));
                string text = element.GetAttribute("title");
                Console.WriteLine("This is a Text: " + text);
                
                WriteToTextFile(text);                   //Function call, write to file

            }
            catch (NoSuchElementException ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

        }

        //function to write cath text to file
        public static void WriteToTextFile(string text)
        {
            try
            {
                //Pass the filepath and filename to the StreamWriter Constructor
                StreamWriter sw = new StreamWriter(@"C:\Users\USER\source\repos\TestTrader_Alex\TextFile1.txt");
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

        //function to wait 10 seconds
        public static void WaitForLoad()
        {
            Stopwatch sw = new Stopwatch();//constructor stop watch
            sw.Start();//start the stopWatch
            for(int i=0; ; i++)
            {
                if(i % 20000 == 0)    
                {
                    sw.Stop();//stop the time measurement
                    if(sw.ElapsedMilliseconds == 10000) // check if  Stopwatch equal 10 seconds
                    {
                        //if time equals 10 seconds stop
                        Console.WriteLine("We waited 10 seconds");
                        break;
                    }
                    else
                    {
                        //if less than 10 seconds
                        sw.Start();
                    }
                }
            }
        }


        public static IWebElement WaitForPageLoad(IWebElement elementBtn)
        {
            int timeout = 1000;
            var sw = new Stopwatch();
            sw.Start();

            while(sw.Elapsed < TimeSpan.FromMilliseconds(timeout))
            {
                try
                {
                    elementBtn = driver.FindElement(By.Id("btnOkLogin"));//to test function use "btnOkLogin1" 
                    break;
                        
                }catch(NoSuchElementException ex)
                {
                    Console.WriteLine(ex);
                }
            }
            sw.Stop();
            return elementBtn;
        }



    }
}
