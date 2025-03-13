using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace UnitTest
{
    [TestClass]
    public class DeleteCart
    {
        [TestMethod]
        public void TestDeleteCart()
        {
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);
            //Đăng nhập vào hệ thống
            IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector(".MuiBox-root.css-1wenov3")));
            loginButton.Click();
            IWebElement emailField = driver.FindElement(By.CssSelector(".MuiOutlinedInput-input.MuiInputBase-input.css-i46v6x-MuiInputBase-input-MuiOutlinedInput-input"));
            emailField.SendKeys("staff");
            IWebElement passwordField = driver.FindElement(By.Name("password"));
            passwordField.SendKeys("123456");
            IWebElement login = driver.FindElement(By.CssSelector(".css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
            login.Click();
            Thread.Sleep(3000);
            IWebElement link = wait.Until(d => d.FindElement(By.CssSelector(".MuiButton-root.MuiButton-text.MuiButton-textPrimary.MuiButton-sizeMedium.MuiButton-textSizeMedium.MuiButtonBase-root.css-170ejer-MuiButtonBase-root-MuiButton-root")));
            link.Click();
            Thread.Sleep(1000);
            IWebElement btnDelete = wait.Until(d => d.FindElements(By.CssSelector(".MuiButtonBase-root.MuiIconButton-root.MuiIconButton-sizeMedium.css-19fss5u-MuiButtonBase-root-MuiIconButton-root")))[1];
            btnDelete.Click();
            Thread.Sleep(3000);
            Console.WriteLine("Delete successfully");
            driver.Quit();
        }


    }
}
