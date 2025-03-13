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
    public class AddCategory
    {
        [TestMethod]
        public void TestAddCategory()
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

            Thread.Sleep(3000);

            // Điều hướng đến trang danh mục
            IWebElement element = driver.FindElement(By.LinkText("Danh mục"));
            element.Click();

            Thread.Sleep(15000);

            // Kiểm tra file dữ liệu
            string filePath = @"C:\DATN\datainputandouput.xlsx";
            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }
            //Tạo workbook từ excel
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(3); // Lấy sheet 4

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        Thread.Sleep(2000);

                        // Thực hiện thêm mới danh mục
                        IWebElement button = driver.FindElement(By.CssSelector(".css-h0hh3m"));
                        button.Click();

                        // Lấy dữ liệu từ Excel
                        string vnl = row.GetCell(0)?.ToString(); // Cột 0: Tên danh mục
                        string name = row.GetCell(1)?.ToString(); // Cột 1: Tiêu đề

                        // Nhập thông tin danh mục
                        IWebElement vnlField = driver.FindElement(By.CssSelector(".css-i46v6x-MuiInputBase-input-MuiOutlinedInput-input"));
                        vnlField.Clear();
                        vnlField.SendKeys(Keys.Control + "a");
                        vnlField.SendKeys(Keys.Delete);
                        vnlField.SendKeys(vnl);

                        IWebElement nameField = driver.FindElement(By.CssSelector(".css-f3h2ax-MuiInputBase-input-MuiOutlinedInput-input"));
                        nameField.Clear();
                        nameField.SendKeys(Keys.Control + "a");
                        nameField.SendKeys(Keys.Delete);
                        nameField.SendKeys(name);

                        // Xác nhận thêm danh mục
                        IWebElement addButton = driver.FindElement(By.CssSelector(".css-75wtpa-MuiDialogActions-root > :not(:first-of-type)"));

                        // Kiểm tra độ dài của tên danh mục
                        Thread.Sleep(1000);
                        int vnlLength = vnlField.GetAttribute("value").Length;
                        int nameLength = nameField.GetAttribute("value").Length;

                        if (vnlLength >= 6 && nameLength >= 1)
                        {
                            row.CreateCell(2).SetCellValue("Pass");
                            addButton.Click();
                        }
                        else if (vnlLength >= 6 && nameLength == 0)
                        {
                            row.CreateCell(2).SetCellValue("Fail_Hệ thống vẫn cho phép để trống trường mô tả");
                            addButton.Click();
                        }
                        else
                        {
                            Console.WriteLine("Tên danh mục không hợp lệ (phải từ 6 đến 25 ký tự).");
                            row.CreateCell(2).SetCellValue("Pass");

                            // Đóng dialog khi nhập không hợp lệ
                            IWebElement cancel = driver.FindElement(By.CssSelector(".css-1bacbjs-MuiButtonBase-root-MuiButton-root"));
                            cancel.Click();
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