using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Threading;

namespace UnitTest
{
    [TestClass]
    public class InputCart
    {
        [TestMethod]
        public void TestInputCart()
        {
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);
            IWebElement loginButton3 = wait.Until(d => d.FindElement(By.CssSelector("li.header__opstion--item.account div.header__opstion--link")));
            loginButton3.Click();
            IWebElement emailField3 = driver.FindElement(By.Id("normal_login_email"));
            emailField3.SendKeys("tranthao102002@gmail.com");
            IWebElement passwordField3 = driver.FindElement(By.Id("normal_login_password"));
            passwordField3.SendKeys("12345678");
            IWebElement login3 = driver.FindElement(By.CssSelector("button[type='submit']"));
            login3.Click();
            Thread.Sleep(5000);

            IList<IWebElement> productDivs = driver.FindElements(By.CssSelector("a.product__content--item"));

            if (productDivs.Count > 0)
            {
                // Chọn phần tử đầu tiên
                //IWebElement firstProductDiv = productDivs[0];
                //((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", firstProductDiv);

                //firstProductDiv.Click();
                Random random = new Random();
                int randomIndex = random.Next(0, 5);
                // Lấy thẻ div ngẫu nhiên từ danh sách
                IWebElement randomProductDiv = productDivs[randomIndex];
                randomProductDiv.Click();
                Thread.Sleep(5000);
                IWebElement detailInfoHead = wait.Until(d => d.FindElement(By.CssSelector("div.detail__btn")));

                IList<IWebElement> buttons = detailInfoHead.FindElements(By.TagName("button"));

                // Kiểm tra nếu danh sách không rỗng
                if (buttons.Count > 0)
                {
                    // Chọn nút đầu tiên
                    IWebElement firstButton = buttons[0];

                    // Cuộn đến nút
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", firstButton);

                    // Nhấn vào nút đầu tiên
                    firstButton.Click();
                    Thread.Sleep(3000);
                    // Tìm nút "OK" bằng CSS Selector
                    IWebElement okButton = wait.Until(d => d.FindElement(By.CssSelector("div.ant-modal-confirm-btns button.ant-btn-primary")));

                    // Nhấp vào nút "OK"
                    okButton.Click();
                    Thread.Sleep(3000);
                    IWebElement cartItem = wait.Until(d => d.FindElement(By.CssSelector("li.header__opstion--item.product__show")));

                    // Nhấp vào thẻ <li>
                    cartItem.Click();
                    IWebElement firstItem = wait.Until(d => d.FindElement(By.CssSelector(".ant-list-items .ant-list-item")));
                    IWebElement inputcartElement = firstItem.FindElement(By.CssSelector(".cart__info--option .detail__info--number--input"));
                    inputcartElement.SendKeys("100");
                    Thread.Sleep(2000);
                    IWebElement notificationMessage = wait.Until(d => d.FindElement(By.XPath("//div[@role='alert']/div[2]")));
                    string messageText = notificationMessage.Text;
                    if (messageText.Length > 0)
                    {
                        Console.WriteLine("Notification Message: " + messageText);
                    }
                    else
                    {
                        IWebElement saveButtonCart = driver.FindElement(By.XPath("//button[@type='button' and span[text()='Lưu']]"));
                        saveButtonCart.Click();
                        Thread.Sleep(4000);
                        IWebElement okButtonCart = wait.Until(d => d.FindElement(By.CssSelector(".ant-btn-primary")));
                        okButtonCart.Click();
                        Thread.Sleep(5000);
                        IList<IWebElement> productItemsinput = driver.FindElements(By.CssSelector("ul.ant-list-items li.ant-list-item"));

                        // Duyệt qua từng sản phẩm và lấy thông tin
                        foreach (IWebElement productItem in productItemsinput)
                        {
                            // Lấy tên sản phẩm
                            IWebElement productNameElement = productItem.FindElement(By.XPath(".//h4[@class='ant-list-item-meta-title']/b"));
                            string productName = productNameElement.Text;

                            // Lấy giá sản phẩm
                            IWebElement productPriceElement = productItem.FindElement(By.XPath(".//b[@class='price']"));
                            string productPrice = productPriceElement.Text;

                            // Lấy số lượng sản phẩm
                            IWebElement productQuantityElement = productItem.FindElement(By.CssSelector(".detail__info--number--input"));
                            string productQuantity = productQuantityElement.GetAttribute("value");

                            // In ra tên, giá và số lượng sản phẩm
                            Console.WriteLine("Product Name: " + productName);
                            Console.WriteLine("Product Price: " + productPrice);
                            Console.WriteLine("Product Quantity: " + productQuantity);
                            Console.WriteLine("----------");
                        }
                    }
                }
            }
            driver.Quit();
        }
    }
}
