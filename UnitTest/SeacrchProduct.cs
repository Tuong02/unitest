using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;

namespace UnitTest
{
    [TestClass]
    public class SeacrchProduct
    {
        [TestMethod]
        public void TestSearchProduct()
        {
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector("li.header__opstion--item.account div.header__opstion--link")));
            loginButton.Click();
            IWebElement emailField = driver.FindElement(By.Id("normal_login_email"));
            emailField.SendKeys("nguyenvancuong13102001t@gmail.com");
            IWebElement passwordField = driver.FindElement(By.Id("normal_login_password"));
            passwordField.SendKeys("1991210a");
            IWebElement login = driver.FindElement(By.CssSelector("button[type='submit']"));
            login.Click();
            Thread.Sleep(4000);
            IWebElement element = driver.FindElement(By.LinkText("Sản phẩm"));
            element.Click();
            Thread.Sleep(3000);
            IWebElement searchInput = driver.FindElement(By.CssSelector("input[placeholder='Tìm theo tên']"));
            searchInput.SendKeys("BÁNH NABATI");
            wait.Until(d => d.FindElement(By.ClassName("ant-table-body")));
            var products = driver.FindElements(By.CssSelector(".ant-table-row"));
            if (products.Count == 0)
            {
                Console.WriteLine("Không có sản phẩm nào.");
            }
            else
            {
                foreach (var product in products)
                {
                    var name = product.FindElement(By.CssSelector("td.ant-table-cell-fix-left")).Text;
                    Console.WriteLine(name);
                }
            }
            driver.Quit();
        }
    }
}
