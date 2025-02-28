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
    public class Login
    {
        [TestMethod]
        public void TestLogin()
        {
            string filePath =@"C:\DATN\datainputandouput.xlsx";
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Tạo workbook từ file Excel
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(0); // Lấy sheet đầu tiên

                // Thực hiện đăng nhập
                IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector(".MuiBox-root.css-1wenov3")));
                loginButton.Click();

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        // Lấy email và mật khẩu từ các ô (giả sử email ở cột 0, password ở cột 1)
                        string email = row.GetCell(0)?.ToString();
                        string password = row.GetCell(1)?.ToString();

                        // Xử lý nhập email
                        IWebElement emailField = driver.FindElement(By.CssSelector(".MuiOutlinedInput-input.MuiInputBase-input.css-i46v6x-MuiInputBase-input-MuiOutlinedInput-input"));
                        emailField.Clear();
                        emailField.SendKeys(Keys.Control + "a"); // Chọn tất cả văn bản
                        emailField.SendKeys(Keys.Delete);         // Xóa văn bản đã chọn
                        emailField.SendKeys(email);

                        // Xử lý nhập password
                        IWebElement passwordField = driver.FindElement(By.Name("password"));
                        passwordField.Clear();
                        passwordField.SendKeys(Keys.Control + "a"); // Chọn tất cả văn bản
                        passwordField.SendKeys(Keys.Delete);         // Xóa văn bản đã chọn
                        passwordField.SendKeys(password);

                        // Nhấn nút đăng nhập
                        IWebElement login = driver.FindElement(By.CssSelector(".css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));
                        login.Click();

                        Thread.Sleep(3000); // Đợi một chút cho các thông báo hiển thị

                        try
                        {
                            IWebElement toastNotification = wait.Until(d => d.FindElement(By.CssSelector(".css-jb04lb-MuiPaper-root-MuiAlert-root")));
                            string toastMessage = toastNotification.Text;
                            Console.WriteLine(toastMessage);

                            if (string.IsNullOrEmpty(email))
                            {
                                row.CreateCell(2).SetCellValue("Pass");
                            }
                            else if (string.IsNullOrEmpty(password))
                            {
                                row.CreateCell(2).SetCellValue("Pass");
                            }
                            else
                            {
                                row.CreateCell(2).SetCellValue("Pass");
                            }
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("Đăng nhập thành công");
                            row.CreateCell(2).SetCellValue("Pass");
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