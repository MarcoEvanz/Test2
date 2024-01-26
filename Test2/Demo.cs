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

namespace Test2
{
    internal class Demo
    {
        IWebDriver driver;
        public static IEnumerable<TestCaseData> GetTestCaseDatasFromExcel()
        {
            var testData = new List<TestCaseData>();
            using (var stream = File.Open("TestCaseData.xlsx", FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var table = result.Tables[0];
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
        public void TestCalcAdd()
        {
            Test();
            driver.FindElement(By.XPath("//input[@id='number1Field']")).SendKeys("11");
            driver.FindElement(By.XPath("//input[@id='number2Field']")).SendKeys("11");
            driver.FindElement(By.XPath("//input[@id='calculateButton']")).Click();
            Thread.Sleep(1000);
            string actual = driver.FindElement(By.XPath("//input[@id='numberAnswerField']")).GetAttribute("value");
            string expected = (11 + 11).ToString();
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
        }
        [Test]
        public void TestCalcSub()
        {
            Test();
            driver.FindElement(By.XPath("//input[@id='number1Field']")).SendKeys("16");
            driver.FindElement(By.XPath("//input[@id='number2Field']")).SendKeys("11");
            driver.FindElement(By.XPath("//option[normalize-space()='Subtract']")).Click();
            driver.FindElement(By.XPath("//input[@id='calculateButton']")).Click();
            Thread.Sleep(1000);
            string actual = driver.FindElement(By.XPath("//input[@id='numberAnswerField']")).GetAttribute("value");
            string expected = (16 - 11).ToString();
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
        }
        [Test]
        public void TestCalcMulti()
        {
            Test();
            driver.FindElement(By.XPath("//input[@id='number1Field']")).SendKeys("16");
            driver.FindElement(By.XPath("//input[@id='number2Field']")).SendKeys("11");
            driver.FindElement(By.XPath("//option[normalize-space()='Multiply']")).Click();
            driver.FindElement(By.XPath("//input[@id='calculateButton']")).Click();
            Thread.Sleep(1000);
            string actual = driver.FindElement(By.XPath("//input[@id='numberAnswerField']")).GetAttribute("value");
            string expected = (16 * 11).ToString();
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
        }
        [Test]
        public void TestCalcDivi()
        {
            Test();
            driver.FindElement(By.XPath("//input[@id='number1Field']")).SendKeys("16");
            driver.FindElement(By.XPath("//input[@id='number2Field']")).SendKeys("11");
            driver.FindElement(By.XPath("//option[normalize-space()='Divide']")).Click();
            driver.FindElement(By.XPath("//input[@id='calculateButton']")).Click();
            Thread.Sleep(1000);
            string actual = driver.FindElement(By.XPath("//input[@id='numberAnswerField']")).GetAttribute("value");
            string expected = (16 / 11).ToString();
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
        }
        [Test]
        [TestCaseSource("GetTestCaseDatasFromExcel")]
        public void TestCalcConca(double a, double b, double expected)
        {
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
        }
        [TearDown]
        public void TearDown()
        {
            driver.Quit();
        }
    }
}
