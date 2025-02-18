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

            //Đăng nhập vào hệ thống
            IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector("li.header__opstion--item.account div.header__opstion--link")));
            loginButton.Click();
            IWebElement emailField = driver.FindElement(By.Id("normal_login_email"));
            emailField.SendKeys("nguyenvancuong13102001t@gmail.com");
            IWebElement passwordField = driver.FindElement(By.Id("normal_login_password"));
            passwordField.SendKeys("1991210a");
            IWebElement login = driver.FindElement(By.CssSelector("button[type='submit']"));
            login.Click();
            Thread.Sleep(4000);
            IWebElement element = driver.FindElement(By.LinkText("Loại sản phẩm"));
            element.Click();
            Thread.Sleep(3000);

            string filePath = @"E:\datainputandouput.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Tạo workbook từ file Excel
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(3); // Lấy sheet 4

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        // Thực hiện thêm mới loại sản phẩm 
                        IWebElement button = driver.FindElement(By.XPath("//button[span[text()='Thêm loại sản phẩm']]"));
                        button.Click();

                        // Lấy tiêu đề và text hiển thị từ các ô
                        string name = row.GetCell(0)?.ToString(); // Giả sử tiêu đề ở cột 0
                        string vnl = row.GetCell(1)?.ToString(); // Giả sử text hiển thị ở cột 1

                        IWebElement nameField = driver.FindElement(By.Id("name"));
                        nameField.Clear(); // Xóa trường nếu có giá trị cũ
                        nameField.SendKeys(name);

                        IWebElement vnlField = driver.FindElement(By.Id("vnl"));
                        vnlField.Clear(); // Xóa trường nếu có giá trị cũ
                        vnlField.SendKeys(vnl);

                        IWebElement addButton = driver.FindElement(By.XPath("//button[span[text()='Thêm mới']]"));
                        addButton.Click();

                        // Kiểm tra kết quả 
                        Thread.Sleep(1000);
                        if (nameField.GetAttribute("value").Length > 0 && vnlField.GetAttribute("value").Length > 0)
                        {
                            if (nameField.GetAttribute("value").Length < 50 && vnlField.GetAttribute("value").Length < 50)
                            {
                                try
                                {
                                    IWebElement notificationMessage = wait.Until(d => d.FindElement(By.XPath("//div[@role='alert']/div[2]")));
                                    string messageText = notificationMessage.Text;
                                    Console.WriteLine("Notification Message: " + messageText);
                                    row.CreateCell(2).SetCellValue("Pass"); // Giả sử ghi vào cột 2
                                }
                                catch (WebDriverTimeoutException)
                                {
                                    Console.WriteLine("Không có toast xuất hiện.");
                                    row.CreateCell(2).SetCellValue("Fail"); // Ghi vào cột 2 nếu không thành công
                                }
                            }
                            else
                            {
                                Console.WriteLine("Hệ thống vẫn cho nhập quá 50 kí tự vào các trường thông tin");
                                row.CreateCell(2).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 50 kí tự vào các trường thông tin"); // Ghi vào cột 2 
                            }
                        }
                        else
                        {
                            Console.WriteLine("Hệ thống vẫn cho để trống các trường thông tin");
                            row.CreateCell(2).SetCellValue("Fail_Hệ thống vẫn cho phép để trống các trường thông tin"); // Ghi vào cột 2 
                        }

                        // Lưu lại file Excel sau khi ghi kết quả
                        using (FileStream writeStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(writeStream);
                        }

                    }
                }
                driver.Quit();
            }
        }
    }
}