using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Assignment
{
    [TestFixture]
    public class TestMethod1: ExcelReader
    {
        ExcelReader x = new ExcelReader();
        string ExpecteTitle = "Yahoo";

   
        [SetUp]
        public void BrowserSetup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("http:www.google.com");
            //System.Threading.Thread.Sleep(5000);
            Console.WriteLine("Browser has opened");
        }

        [Test]
        public void TestMethod()
        {
            for (int i = 1; i < 5; i++)
            {

                driver.FindElement(By.XPath("//input[@id='lst-ib']")).SendKeys(x.ExcelDataReader(i, 1));
                driver.FindElement(By.XPath("//input[@id='lst-ib']")).SendKeys(Keys.Enter);
                

                if(x.ExcelDataReader(i, 1).Equals("Yahoo"))
                {

                    driver.FindElement(By.XPath("//cite[contains(text(),'https://sg.yahoo.com/')]")).Click();
                    String Ytitle = driver.Title;
                    Console.WriteLine("The Title of page:-"+Ytitle);
                    Console.WriteLine("Test Is Passed");
                    driver.Navigate().Back();
                    driver.FindElement(By.XPath("//input[@id='lst-ib']")).Clear();
                    System.Threading.Thread.Sleep(5000);
                    continue;
                }
                
                
                driver.FindElement(By.XPath("//input[@id='lst-ib']")).Clear();
            }
        }

        [TearDown]
        public void TearDown()
        {
            driver.Close();
        }
    }
}
