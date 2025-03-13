using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Threading;

namespace UnitTest
{
    [TestClass]
    public class DeleteBrand
    {
        [TestMethod]
        public void TestDeleteBrand()
        {
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);

            // Đăng nhập vào hệ thống
            IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector(".MuiBox-root.css-1wenov3")));
            loginButton.Click();
            IWebElement emailField = driver.FindElement(By.CssSelector(".MuiOutlinedInput-input.MuiInputBase-input.css-i46v6x-MuiInputBase-input-MuiOutlinedInput-input"));
            emailField.SendKeys("admin");
            IWebElement passwordField = driver.FindElement(By.Name("password"));
            passwordField.SendKeys("@dmIn12");
            IWebElement login = driver.FindElement(By.CssSelector(".css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
            login.Click();
            Thread.Sleep(10000);

            // click button brand
            IWebElement element = wait.Until(d => d.FindElement(By.LinkText("Thương hiệu")));
            element.Click();
            Thread.Sleep(5000);

            IWebElement chose = driver.FindElements(By.CssSelector(".MuiButtonBase-root.MuiIconButton-root.MuiIconButton-sizeMedium.css-19fss5u-MuiButtonBase-root-MuiIconButton-root"))[4];
            chose.Click();
            Thread.Sleep(500);

            IWebElement deleteButton = driver.FindElement(By.XPath("//span[text()='Xóa']"));
            deleteButton.Click();
            Thread.Sleep(1000);

            IWebElement saveButton = driver.FindElement(By.XPath("//button[span[text()='Lưu']]"));
            saveButton.Click();
            Thread.Sleep(500);

            Console.WriteLine("Xóa thành công");

            driver.Quit();
        }
    }
}