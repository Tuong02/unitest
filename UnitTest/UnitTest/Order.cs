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

            // Đăng nhập vào hệ thống
            IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector(".MuiBox-root.css-1wenov3")));
            loginButton.Click();
            IWebElement emailField = driver.FindElement(By.CssSelector(".MuiOutlinedInput-input.MuiInputBase-input.css-i46v6x-MuiInputBase-input-MuiOutlinedInput-input"));
            emailField.SendKeys("staff");
            IWebElement passwordField = driver.FindElement(By.Name("password"));
            passwordField.SendKeys("123456");
            IWebElement login = driver.FindElement(By.CssSelector(".css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
            login.Click();
            Thread.Sleep(4000);

            string filePath = @"C:\DATN\datainputandouput.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Tạo workbook từ file Excel
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(10); // Lấy sheet 5

                Thread.Sleep(5000);
                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        try
                        {
                            string products = row.GetCell(0)?.ToString();

                            // Điều hướng đến danh sách sản phẩm
                            IWebElement btnSearch = wait.Until(d => d.FindElement(By.CssSelector(".css-tihntr-MuiInputBase-input-MuiOutlinedInput-input")));
                            btnSearch.SendKeys(Keys.Control + "a");
                            btnSearch.SendKeys(products);
                            btnSearch.SendKeys(Keys.Enter);
                            Thread.Sleep(7000);

                            string xpath = $"//a[@aria-label='{products}']";
                            IWebElement btnChose = wait.Until(d => d.FindElement(By.XPath(xpath)));
                            btnChose.Click();
                            Thread.Sleep(3000);

                            IWebElement muaNgayButton = wait.Until(d => d.FindElement(By.XPath("//button[span[text()='Mua ngay']]")));
                            muaNgayButton.Click();
                            Thread.Sleep(5000);

                            IWebElement btnPay = wait.Until(d => d.FindElement(By.XPath("//button[span[text()='Thanh toán']]")));
                            btnPay.Click();
                            Thread.Sleep(3000);

                            IWebElement radioButton = wait.Until(d => d.FindElement(By.CssSelector("input[name='userAddressId']")));
                            radioButton.Click();
                            Thread.Sleep(3000);

                            IWebElement continueButton = wait.Until(d => d.FindElement(By.XPath("//button[span[text()='Tiếp tục']]")));
                            continueButton.Click();
                            Thread.Sleep(3000);

                            IWebElement codRadio = wait.Until(d => d.FindElement(By.XPath("//input[@name='paymentMethod'][@value='cod']")));
                            codRadio.Click();
                            Thread.Sleep(3000);

                            IWebElement orderButton = wait.Until(d => d.FindElement(By.XPath("//button[span[text()='Đặt hàng']]")));
                            orderButton.Click();
                            Thread.Sleep(5000);

                            row.CreateCell(1).SetCellValue("Pass_Đặt hàng thành công");
                        }
                        catch (WebDriverTimeoutException)
                        {
                            Console.WriteLine("Không có toast xuất hiện.");
                            row.CreateCell(1).SetCellValue("fail");
                        }
                    }
                }

                // Lưu lại file Excel sau khi ghi kết quả
                using (FileStream writeStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(writeStream);
                }
            }

            driver.Quit();
        }
    }
}
