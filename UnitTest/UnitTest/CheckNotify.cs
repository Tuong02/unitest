using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace UnitTest
{
    [TestClass]
    public class CheckNotify
    {
        [TestMethod]
        public void TestCheckNotify()
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
            IWebElement element = driver.FindElement(By.LinkText("Thông báo"));
            element.Click();
            Thread.Sleep(3000);
            IWebElement firstSendAgainButton = driver.FindElement(By.XPath("//ul[@class='ant-list-items']/li[1]//button[span[text()='Gửi lại']]"));
            firstSendAgainButton.Click();
            Thread.Sleep(2000);
            IWebElement okButton2 = driver.FindElement(By.XPath("//button[span[text()='OK']]"));
            okButton2.Click();
            Actions actions = new Actions(driver);
            IWebElement adminDiv = driver.FindElement(By.XPath("//div[contains(@class, 'ant-dropdown-trigger')]//p[text()='Admin']"));
            adminDiv.Click();
            wait.Until(d => d.FindElement(By.XPath("//span[text()='Đăng xuất']")));
            IWebElement logoutButton = driver.FindElement(By.XPath("//span[text()='Đăng xuất']"));
            logoutButton.Click();
            IWebElement okButtonlogout = driver.FindElement(By.XPath("//button[span='OK']"));
            okButtonlogout.Click();
            Thread.Sleep(4000);
            IWebElement loginButtonNotification = wait.Until(d => d.FindElement(By.CssSelector("li.header__opstion--item.account div.header__opstion--link")));
            loginButtonNotification.Click();
            IWebElement emailNotification = driver.FindElement(By.Id("normal_login_email"));
            emailNotification.SendKeys("tranthao102002@gmail.com");
            IWebElement passwordNotification = driver.FindElement(By.Id("normal_login_password"));
            passwordNotification.SendKeys("12345678");
            IWebElement loginNotification = driver.FindElement(By.CssSelector("button[type='submit']"));
            loginNotification.Click();
            Thread.Sleep(4000);
            IWebElement notificationElement = driver.FindElement(By.XPath("//li[contains(@class, 'shop') and .//p[text()='Thông báo']]"));
            notificationElement.Click();
            wait.Until(d => d.FindElement(By.CssSelector("div.wrap_notify")));
            IList<IWebElement> notifications = driver.FindElements(By.CssSelector("div.wrap_notify .items.unread__notify"));
            // Lấy 5 thông báo mới nhất
            var latestNotifications = notifications.Take(5);
            foreach (var notification in latestNotifications)
            {
                string titleNotification = notification.FindElement(By.CssSelector("div.content b")).Text;
                string message = notification.FindElement(By.CssSelector("div.content .price")).Text;
                string timeAgo = notification.FindElement(By.XPath(".//div[2]")).Text;

                Console.WriteLine($"Title: {titleNotification}");
                Console.WriteLine($"Message: {message}");
                Console.WriteLine($"Time: {timeAgo}");
                Console.WriteLine("=====================");
            }

            driver.Quit();
        }
    }
}
