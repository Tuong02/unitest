using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace UnitTest
{
    [TestClass]
    public class Order
    {
        [TestMethod]
        public void TestOrder()
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
            string filePath = @"E:\datainputandouput.xlsx";


            IList<IWebElement> productDivs = driver.FindElements(By.CssSelector("a.product__content--item"));

            if (productDivs.Count > 0)
            {
                // Chọn phần tử đầu tiên
                //IWebElement firstProductDiv = productDivs[0];
                //((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", firstProductDiv);
                //firstProductDiv.Click();
                // random 
                Random random = new Random();
                int randomIndex = random.Next(0, 5);

                // Lấy thẻ div ngẫu nhiên từ danh sách
                IWebElement randomProductDiv = productDivs[randomIndex];
                randomProductDiv.Click();
                Thread.Sleep(4000);
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
                    Thread.Sleep(5000);
                    IList<IWebElement> checkboxes = driver.FindElements(By.CssSelector(".ant-checkbox-input"));

                    foreach (IWebElement checkbox in checkboxes)
                    {
                        checkbox.Click();
                    }
                    Thread.Sleep(5000);
                    // Tìm nút "Đặt hàng" 
                    IWebElement orderButton = driver.FindElement(By.XPath("//button[span[text()='Đặt hàng']]"));

                    // Thực hiện click vào nút "Đặt hàng"
                    orderButton.Click();
                    Thread.Sleep(4000);
                    //IWebElement phoneInput = driver.FindElement(By.XPath("//input[@placeholder='Số điện thoại']"));
                    //phoneInput.Clear();
                    //phoneInput.SendKeys("092456789");
                    //IWebElement addressInput = driver.FindElement(By.XPath("//input[@placeholder='Địa chỉ giao hàng']"));
                    //addressInput.Clear();
                    //addressInput.SendKeys("Nam Định");
                    IWebElement shippingCostButton = driver.FindElement(By.XPath("//button[span[text()='Giá ship']]"));
                    shippingCostButton.Click();
                    Thread.Sleep(4000);
                    IWebElement oder = wait.Until(d => d.FindElement(By.ClassName("billing-button")));

                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine("File không tồn tại.");
                        return;
                    }

                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {


                        // Tạo workbook từ file Excel
                        IWorkbook workbook = new XSSFWorkbook(fileStream);
                        ISheet sheet = workbook.GetSheetAt(9); // Lấy sheet 5
                        oder.Click();
                        Thread.Sleep(5000);
                        for (int i = 1; i <= 1; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                        {
                            IRow row = sheet.GetRow(i);
                            if (row != null)
                            {
                                try
                                {
                                    IWebElement toastNotification = wait.Until(d => d.FindElement(By.CssSelector("div.Toastify__toast-body")));
                                    string toastMessage = toastNotification.Text;
                                    Console.WriteLine("Thông báo từ toast: " + toastMessage);
                                    row.CreateCell(1).SetCellValue("Pass");

                                }
                                catch (WebDriverTimeoutException)
                                {
                                    Console.WriteLine("Không có toast xuất hiện.");
                                    row.CreateCell(1).SetCellValue("fail");
                                }
                            }
                        }

                        using (FileStream writeStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(writeStream);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Không tìm thấy nút nào.");
                }
            }
            else
            {
                Console.WriteLine("Không tìm thấy sản phẩm nào.");
            }
            driver.Quit();
        }
    }
}
