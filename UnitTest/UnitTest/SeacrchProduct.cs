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
    public class SeacrchProduct
    {
        [TestMethod]
        public void TestSearchProduct()
        {
            string filePath = @"C:\DATN\datainputandouput.xlsx";
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }

            // Đăng nhập vào hệ thống
            IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector(".MuiBox-root.css-1wenov3")));
            loginButton.Click();

            IWebElement emailField = driver.FindElement(By.CssSelector(".MuiOutlinedInput-input.MuiInputBase-input.css-i46v6x-MuiInputBase-input-MuiOutlinedInput-input"));
            emailField.SendKeys("admin");

            IWebElement passwordField = driver.FindElement(By.Name("password"));
            passwordField.SendKeys("@dmIn12");

            IWebElement login = driver.FindElement(By.CssSelector(".css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
            login.Click();

            Thread.Sleep(4000);

            // Điều hướng đến danh sách sản phẩm
            IWebElement element = driver.FindElement(By.XPath("//div[@role='button' and .//div[text()='Sản phẩm']]"));
            element.Click();
            Thread.Sleep(500);

            IWebElement listProduct = driver.FindElement(By.CssSelector("a[href='/dashboard/app/products/list']"));
            listProduct.Click();
            Thread.Sleep(15000);

            // Mở file Excel
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(9); // Lấy sheet 10

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        string productName = row.GetCell(0)?.ToString(); // Giả sử tên sản phẩm ở cột 0

                        // Nhập vào ô tìm kiếm sản phẩm
                        IWebElement inputProduct = driver.FindElement(By.CssSelector(".css-tihntr-MuiInputBase-input-MuiOutlinedInput-input"));
                        inputProduct.Clear();
                        inputProduct.SendKeys(Keys.Control + "a");
                        inputProduct.SendKeys(productName);
                        inputProduct.SendKeys(Keys.Enter);

                        Thread.Sleep(7000);

                        try
                        {
                            // Kiểm tra xem sản phẩm có hiển thị hay không
                            string result = driver.FindElement(By.CssSelector(".MuiTypography-root.MuiTypography-h4.MuiTypography-paragraph.css-85n264-MuiTypography-root")).Text;
                            if (result.Contains("Không tìm thấy sản phẩm"))
                            {
                                row.CreateCell(1).SetCellValue("Không tìm thấy sản phẩm");
                            } 
                            
                        }
                        catch (NoSuchElementException)
                        {
                            // Nếu không tìm thấy thông báo lỗi, tức là sản phẩm đã hiển thị
                            row.CreateCell(1).SetCellValue("Pass");
                        }
                        catch (WebDriverTimeoutException)
                        {
                            row.CreateCell(1).SetCellValue("Đã xảy ra lỗi khi tìm kiếm");
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
