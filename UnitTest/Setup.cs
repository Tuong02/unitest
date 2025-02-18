using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;

namespace UnitTest
{
    public class Setup
    {
        public IWebDriver SetupChromeDriver()
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:4200/home");
            return driver;
        }

        public WebDriverWait CreateWebDriverWait(IWebDriver driver)
        {
            return new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        }
    }
}
