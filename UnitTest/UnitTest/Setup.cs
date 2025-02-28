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
            new WebDriverManager.DriverManager().SetUpDriver(new WebDriverManager.DriverConfigs.Impl.ChromeConfig());
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://hk-ecommerce.vercel.app/");
            return driver;
        }

        public WebDriverWait CreateWebDriverWait(IWebDriver driver)
        {
            return new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        }
    }
}
