using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using ExcelDataReader;
using OfficeOpenXml;
using NUnit.Framework.Interfaces;

namespace Test2
{
    internal class Demo
    {
        IWebDriver driver;
        public static IEnumerable<TestCaseData> GetTestCaseDatasFromExcel(string sheetName)
        {
            var testData = new List<TestCaseData>();
            using (var stream = File.Open("TestCaseData.xlsx", FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var table = result.Tables[sheetName];
                    for (int i= 0; i < table.Rows.Count; i++)
                    {
                        double num1 = Convert.ToDouble(table.Rows[i][0].ToString());
                        double num2 = Convert.ToDouble(table.Rows[i][1].ToString());
                        double expected = Convert.ToDouble(table.Rows[i][2].ToString());
                        testData.Add(new TestCaseData(num1, num2, expected));
                    }
                }
            }
            return testData;
        }
        [SetUp]
        public void Setup()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            ChromeDriverService service = ChromeDriverService.CreateDefaultService("D:\\chromedriver-win64");
            driver = new ChromeDriver(service, options);


        }
        [Test]
        public void Test()
        {
            driver.Navigate().GoToUrl("https://testsheepnz.github.io/BasicCalculator.html");
            Thread.Sleep(2000);
        }

        [Test]
        [TestCaseSource(nameof(GetTestCaseDatasFromExcel), new object[] { "Sub" })]
        public void TestCalcAdd(double a, double b, double expected)
        {
            string sheetname = "Add";
            Test();
            string num1 = a.ToString();
            string num2 = b.ToString();
            driver.FindElement(By.XPath("//input[@id='number1Field']")).SendKeys(num1);
            driver.FindElement(By.XPath("//input[@id='number2Field']")).SendKeys(num2);
            driver.FindElement(By.XPath("//input[@id='calculateButton']")).Click();
            Thread.Sleep(1000);
            string result = driver.FindElement(By.XPath("//input[@id='numberAnswerField']")).GetAttribute("value");
            double actual = Convert.ToDouble(result);
            //if (actual != Convert.ToString(expected))
            //{
            //    Console.WriteLine(actual);
            //    Console.WriteLine(expected);
            //    Console.WriteLine("False");
            //}
            if (actual != expected)
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("False");
            }
            else
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("True");
            }
            WriteDataToExcel(Convert.ToString(actual), sheetname);
        }
        [Test]
        [TestCaseSource(nameof(GetTestCaseDatasFromExcel), new object[] { "Sub" })]
        public void TestCalcSub(double a, double b, double expected)
        {
            string sheetname = "Sub";
            Test();
            string num1 = a.ToString();
            string num2 = b.ToString();
            driver.FindElement(By.XPath("//input[@id='number1Field']")).SendKeys(num1);
            driver.FindElement(By.XPath("//input[@id='number2Field']")).SendKeys(num2);
            driver.FindElement(By.XPath("//option[normalize-space()='Subtract']")).Click();
            driver.FindElement(By.XPath("//input[@id='calculateButton']")).Click();
            Thread.Sleep(1000);
            string result = driver.FindElement(By.XPath("//input[@id='numberAnswerField']")).GetAttribute("value");
            double actual = Convert.ToDouble(result);
            if (actual != expected)
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("False");
            }
            else
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("True");
            }
            WriteDataToExcel(Convert.ToString(actual), sheetname);
        }
        [Test]
        [TestCaseSource(nameof(GetTestCaseDatasFromExcel), new object[] { "Multi" })]
        public void TestCalcMulti(double a, double b, double expected)
        {
            string sheetname = "Multi";
            Test();
            string num1 = a.ToString();
            string num2 = b.ToString();
            driver.FindElement(By.XPath("//input[@id='number1Field']")).SendKeys(num1);
            driver.FindElement(By.XPath("//input[@id='number2Field']")).SendKeys(num2);
            driver.FindElement(By.XPath("//option[normalize-space()='Multiply']")).Click();
            driver.FindElement(By.XPath("//input[@id='calculateButton']")).Click();
            Thread.Sleep(1000);
            string result = driver.FindElement(By.XPath("//input[@id='numberAnswerField']")).GetAttribute("value");
            double actual = Convert.ToDouble(result);
            if (actual != expected)
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("False");
            }
            else
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("True");
            }
            WriteDataToExcel(Convert.ToString(actual), sheetname);
        }
        [Test]
        [TestCaseSource(nameof(GetTestCaseDatasFromExcel), new object[] { "Divi" })]
        public void TestCalcDivi(double a, double b, double expected)
        {
            string sheetname = "Divi";
            Test();
            string num1 = a.ToString();
            string num2 = b.ToString();
            driver.FindElement(By.XPath("//input[@id='number1Field']")).SendKeys(num1);
            driver.FindElement(By.XPath("//input[@id='number2Field']")).SendKeys(num2);
            driver.FindElement(By.XPath("//option[normalize-space()='Divide']")).Click();
            driver.FindElement(By.XPath("//input[@id='calculateButton']")).Click();
            Thread.Sleep(1000);
            string result = driver.FindElement(By.XPath("//input[@id='numberAnswerField']")).GetAttribute("value");
            double actual = Convert.ToDouble(result);
            if (actual != expected)
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("False");
            }
            else
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("True");
            }
            WriteDataToExcel(Convert.ToString(actual), sheetname);
        }
        [Test]
        [TestCaseSource(nameof(GetTestCaseDatasFromExcel), new object[] { "Conca" })]

        public void TestCalcConca(double a, double b, double expected)
        {
            string sheetname = "Conca";
            Test();
            string num1 = a.ToString();
            string num2 = b.ToString();
            driver.FindElement(By.XPath("//input[@id='number1Field']")).SendKeys(num1);
            driver.FindElement(By.XPath("//input[@id='number2Field']")).SendKeys(num2);
            driver.FindElement(By.XPath("//option[normalize-space()='Concatenate']")).Click();
            driver.FindElement(By.XPath("//input[@id='calculateButton']")).Click();
            Thread.Sleep(1000);
            string result = driver.FindElement(By.XPath("//input[@id='numberAnswerField']")).GetAttribute("value");
            double actual = Convert.ToDouble(result);
            if (actual != expected)
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("False");
            }
            else
            {
                Console.WriteLine(actual);
                Console.WriteLine(expected);
                Console.WriteLine("True");
            }
            WriteDataToExcel(Convert.ToString(actual), sheetname);
        }
        [TearDown]
        public void TearDown()
        {
            
            driver.Quit();
        }
        public void WriteDataToExcel(String Actual, String SheetName)
        {
            // Đường dẫn của tệp Excel đích
            string excelFilePath = "TestCaseData.xlsx";

            // Tạo một tệp Excel mới
            using (var excelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                // Lấy hoặc tạo một Sheet có tên được truyền vào
                var worksheet = excelPackage.Workbook.Worksheets[SheetName];

                // Ghi dữ liệu vào các ô trong Sheet
                int lastRow = 1;
                while (worksheet.Cells[lastRow, 4].Value != null)
                {
                    lastRow++;
                }

                // Ghi dữ liệu vào ô ở dòng mới sau dòng cuối cùng
                worksheet.Cells[lastRow, 4].Value = Actual;

                // Lưu tệp Excel
                excelPackage.Save();
            }

            Console.WriteLine("Data has been written to Excel successfully.");
        }
        //public static void WriteDataToExcel(IEnumerable<TestCaseData> testData, string excelFilePath)
        //{
        //    using (var excelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
        //    {
        //        var worksheet = excelPackage.Workbook.Worksheets["Data"] ?? excelPackage.Workbook.Worksheets.Add("Data");

        //        int row = 1;
        //        foreach (var data in testData)
        //        {
        //            worksheet.Cells[row, 1].Value = data.Arguments[0];
        //            worksheet.Cells[row, 2].Value = data.Arguments[1];
        //            worksheet.Cells[row, 3].Value = data.Arguments[2];
        //            row++;
        //        }

        //        excelPackage.Save();
        //    }

        //    Console.WriteLine("Data has been written to Excel successfully.");
        //}
    }
}
