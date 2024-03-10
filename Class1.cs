using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.Events;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Security.Policy;
using OpenQA.Selenium.DevTools.V118.Browser;


namespace Robo
{
    internal class Class1
    {
       
        static void Main()
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            var chromeDriverPath = @"C:\Users\thoma\OneDrive\Área de Trabalho\Nova pasta\chromedriver-win64";
            var chromeOptions = new ChromeOptions();


            IWebDriver driver = new ChromeDriver(chromeDriverPath, chromeOptions);

            driver.Navigate().GoToUrl("https://rpachallenge.com/");

            var entrar = driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button"));
            entrar.Click();

            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\thoma\\OneDrive\\Área de Trabalho\\Robo\\TreinamentoRPA\\desafio.xlsx"))) 

            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            
            int numeroDePessoas = worksheet.Dimension.Rows;

                for (int i = 2; i <= 11; i++)
                {
                    string firstName = worksheet.Cells[i, 1].Text;
                    string lastName = worksheet.Cells[i, 2].Text;
                    string companyName = worksheet.Cells[i, 3].Text;
                    string roleInCompany = worksheet.Cells[i, 4].Text;
                    string address = worksheet.Cells[i, 5].Text;
                    string email = worksheet.Cells[i, 6].Text;
                    string phoneNumber = worksheet.Cells[i, 7].Text;

                    driver.FindElement(By.XPath("//*[@ng-reflect-name='labelFirstName']")).SendKeys(firstName);
                    Thread.Sleep(100);

                    driver.FindElement(By.XPath("//*[@ng-reflect-name='labelLastName']")).SendKeys(lastName);
                    Thread.Sleep(100);

                    driver.FindElement(By.XPath("//*[@ng-reflect-name='labelCompanyName']")).SendKeys(companyName);
                    Thread.Sleep(100);

                    driver.FindElement(By.XPath("//*[@ng-reflect-name='labelRole']")).SendKeys(roleInCompany);
                    Thread.Sleep(100);

                    driver.FindElement(By.XPath("//*[@ng-reflect-name='labelAddress']")).SendKeys(address);
                    Thread.Sleep(100);

                    driver.FindElement(By.XPath("//*[@ng-reflect-name='labelEmail']")).SendKeys(email);
                    Thread.Sleep(100);

                    driver.FindElement(By.XPath("//*[@ng-reflect-name='labelPhone']")).SendKeys(phoneNumber);
                    Thread.Sleep(100);

                    driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input")).Click();
                    Thread.Sleep(1000);
                }
            }

            Thread.Sleep(500);


            Console.ReadLine();
            driver.Quit();
        }

    }
}
