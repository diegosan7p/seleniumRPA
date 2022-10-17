using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Data.OleDb;
namespace SeleniumTest
{
    class Program
    {


        public DataTable getDTFromExcel(string path)
        {

            OleDbConnection connection = null;
            OleDbCommand cmd = null;
            OleDbDataAdapter adapter = null;
            DataTable result = null;
            DataSet dataSet = null;

            try
            {
                connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                                           "Data Source='" + path + "';" +
                                           "Extended Properties='Excel 12.0 Xml;HDR=YES'");
                cmd = new OleDbCommand();
                cmd.Connection = connection;
                cmd.CommandText = "select * from [Sheet1$]";
                adapter = new OleDbDataAdapter(cmd);
                dataSet = new DataSet();
                adapter.Fill(dataSet);
                result = new DataTable();
                result = dataSet.Tables[0];
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                if(connection!=null)
                {
                    connection.Dispose();
                }
                
            }
            return result;
        }


        static void Main(string[] args)
        {

            string path = @"C:\Users\diego\Downloads\challenge.xlsx";
            Program p = new Program();

            DataTable extractedDT = p.getDTFromExcel(path);

            IWebDriver chromeDriver = new ChromeDriver(Environment.CurrentDirectory);

            chromeDriver.Navigate().GoToUrl("http://www.rpachallenge.com/");

            chromeDriver.Manage().Window.Maximize();

            chromeDriver.FindElement(By.XPath("//button[text()= 'Start']")).Click();

            foreach (DataRow row in extractedDT.Rows)
            {
                //first name index 0
                chromeDriver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelFirstName']")).SendKeys(row[0].ToString());
                //last name index 1
                chromeDriver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelLastName']")).SendKeys(row[1].ToString());
                //company name index 
                chromeDriver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelCompanyName']")).SendKeys(row[2].ToString());
                //Role In Comp index 3
                chromeDriver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelRole']")).SendKeys(row[3].ToString());
                //address index 4
                chromeDriver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelAddress']")).SendKeys(row[4].ToString());
                //email index 5
                chromeDriver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelEmail']")).SendKeys(row[5].ToString());
                //phone index 6
                chromeDriver.FindElement(By.XPath("//input[@ng-reflect-name= 'labelPhone']")).SendKeys(row[6].ToString());

                chromeDriver.FindElement(By.XPath("//input[@value= 'Submit']")).Click();

            }
            string rate = chromeDriver.FindElement(By.XPath("//div[@class= 'message2']")).GetAttribute("value");

            
            Console.ReadLine();


        }
    }
}
