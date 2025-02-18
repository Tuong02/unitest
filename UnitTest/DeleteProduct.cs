using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;

namespace UnitTest
{
    [TestClass]
    public class DeleteProduct
    {
        [TestMethod]
        public void TestDeleteProduc()
        {
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);
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
            IWebElement firstRowdelete = wait.Until(d => d.FindElement(By.XPath("//tbody/tr[1]")));
            IWebElement detailButtondelete = driver.FindElement(By.XPath(".//button[contains(., 'Xóa')]"));
            detailButtondelete.Click(); ;
            Thread.Sleep(2000);
            IWebElement okButton = driver.FindElement(By.XPath("//button[span[text()='OK']]"));
            okButton.Click();
            Thread.Sleep(1000);
            try
            {
                IWebElement notificationMessage = wait.Until(d => d.FindElement(By.XPath("//div[@role='alert']/div[2]")));
                string messageText = notificationMessage.Text;
                Console.WriteLine("Notification Message: " + messageText);
            }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Không có toast xuất hiện.");
            }
        }
    }
}
