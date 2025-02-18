using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Threading;

namespace UnitTest
{
    [TestClass]
    public class UpdateAccount
    {
        [TestMethod]
        public void TestUpdateAccount()
        {
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);

            string filePath = @"E:\datainputandouput.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }
            //Đăng nhập vào hệ thống 
            IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector("li.header__opstion--item.account div.header__opstion--link")));
            loginButton.Click();
            IWebElement emailField = wait.Until(d => d.FindElement(By.Id("normal_login_email")));
            emailField.SendKeys("nguyenvancuong13102001t@gmail.com");
            IWebElement passwordField = wait.Until(d => d.FindElement(By.Id("normal_login_password")));
            passwordField.SendKeys("1991210a");
            IWebElement login = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
            login.Click();
            Thread.Sleep(4000);
            IWebElement element = wait.Until(d => d.FindElement(By.LinkText("Tài khoản")));
            element.Click();
            Thread.Sleep(3000);


            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Tạo workbook từ file Excel
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(5); // Lấy sheet 6

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        string username = row.GetCell(0)?.ToString(); // Giả sử tiêu đề ở cột 0
                        string age = row.GetCell(1).ToString(); // Giả sử text hiển thị ở cột 1
                        string email = row.GetCell(2)?.ToString(); // Giả sử text hiển thị ở cột 2
                        string address = row.GetCell(3)?.ToString(); // Giả sử text hiển thị ở cột 3
                        string phoneNumber = row.GetCell(4)?.ToString(); // Giả sử text hiển thị ở cột 4
                        IWebElement twoRow = wait.Until(d => d.FindElement(By.XPath("//tbody/tr[2]")));
                        //Thực hiện update tài khoản 
                        IWebElement detailButtonTwo = wait.Until(d => d.FindElement(By.XPath(".//button[contains(., 'Chi tiết')]")));
                        detailButtonTwo.Click();
                        Thread.Sleep(2000);
                        IWebElement usernameField = wait.Until(d => d.FindElement(By.Id("username")));
                        usernameField.Clear(); // Xóa trường nếu có giá trị cũ
                        usernameField.SendKeys(username);

                        IWebElement ageField = wait.Until(d => d.FindElement(By.Name("age")));
                        ageField.Clear(); // Xóa trường nếu có giá trị cũ
                        ageField.SendKeys(age);

                        IWebElement emaillField = wait.Until(d => d.FindElement(By.Id("email")));
                        emaillField.Clear(); // Xóa trường nếu có giá trị cũ
                        emaillField.SendKeys(email);

                        IWebElement addressField = wait.Until(d => d.FindElement(By.Id("address")));
                        addressField.Clear(); // Xóa trường nếu có giá trị cũ
                        addressField.SendKeys(address);

                        IWebElement phoneNumberField = wait.Until(d => d.FindElement(By.Id("phoneNumber")));
                        phoneNumberField.Clear(); // Xóa trường nếu có giá trị cũ
                        phoneNumberField.SendKeys(phoneNumber);

                        IWebElement addButton = wait.Until(d => d.FindElement(By.XPath("//button[span[text()='Update']]")));
                        addButton.Click();

                        string emailPattern = @"^[^@\s]+@[^@\s]+\.[^@\s]+$";

                        // Kiểm tra kết quả 
                        Thread.Sleep(1000);
                        if (usernameField.GetAttribute("value").Length > 0 && ageField.GetAttribute("value").Length > 0 && emaillField.GetAttribute("value").Length > 0 && addressField.GetAttribute("value").Length > 0 && phoneNumberField.GetAttribute("value").Length > 0)
                        {
                            if (usernameField.GetAttribute("value").Length < 50 && addressField.GetAttribute("value").Length < 50)
                            {
                                if (!System.Text.RegularExpressions.Regex.IsMatch(email, emailPattern))
                                {
                                    Console.WriteLine("Hệ thống vẫn cho phép nhập không đúng định dạng email");
                                    row.CreateCell(5).SetCellValue("Fail_Hệ thống vẫn cho phép nhập không đúng định dạng email"); // Giả sử ghi vào cột 2
                                }
                                else if (!System.Text.RegularExpressions.Regex.IsMatch(age, @"^\d+$"))
                                {
                                    Console.WriteLine("Hệ thống vẫn cho phép nhập Không đúng định dạng tuổi");
                                    row.CreateCell(5).SetCellValue("Fail_Hệ thống vẫn cho phép nhập không đúng định dạng tuổi"); // Giả sử ghi vào cột 2
                                }
                                else if (!System.Text.RegularExpressions.Regex.IsMatch(phoneNumber, @"^\d{10}$"))
                                {
                                    Console.WriteLine("Hệ thống vẫn cho phép nhập không đúng định dạng số điện thoại");
                                    row.CreateCell(5).SetCellValue("Fail_Hệ thống vẫn cho phép nhập không đúng định dạng số điện thoại"); // Giả sử ghi vào cột 2
                                }
                                else
                                {
                                    try
                                    {
                                        IWebElement notificationMessage = wait.Until(d => d.FindElement(By.XPath("//div[@role='alert']/div[2]")));
                                        string messageText = notificationMessage.Text;
                                        Console.WriteLine("Notification Message: " + messageText);
                                        row.CreateCell(5).SetCellValue("Pass"); // Giả sử ghi vào cột 2
                                    }
                                    catch (WebDriverTimeoutException)
                                    {
                                        Console.WriteLine("Không có toast xuất hiện.");
                                        row.CreateCell(5).SetCellValue("Fail"); // Giả sử ghi vào cột 2
                                    }

                                }
                            }
                            else
                            {
                                Console.WriteLine("Hệ thống vẫn cho nhập quá 50 kí tự vào các trường thông tin");
                                row.CreateCell(5).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 50 kí tự vào các trường thông tin"); // Ghi vào cột 2 }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Hệ thống vẫn cho để trống các trường thông tin");
                            row.CreateCell(5).SetCellValue("Fail_Hệ thống vẫn cho phép để trống các trường thông tin"); // Ghi vào cột 2 
                        }

                        // Lưu lại file Excel sau khi ghi kết quả
                        using (FileStream writeStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(writeStream);
                        }

                    }
                    else
                    {
                        driver.Quit();
                    }
                }
                driver.Quit();

            }
        }
    }
}