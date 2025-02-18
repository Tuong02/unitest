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
            string filePath = @"C:\DATN\datainputandouput.xlsx"; 
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
                IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector("menu--product-categories")));
                loginButton.Click();

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        // Lấy email và mật khẩu từ các ô

                        string email = row.GetCell(0)?.ToString(); // Giả sử email ở cột 0
                        string password = row.GetCell(1)?.ToString(); // Giả sử mật khẩu ở cột 1


                        IWebElement emailField = driver.FindElement(By.Id("form-control ng-pristine ng-valid ng-touched"));
                        emailField.Clear(); // Xóa trường nếu có giá trị cũ
                        emailField.SendKeys(email);

                        IWebElement passwordField = driver.FindElement(By.Id("form-control ng-untouched ng-pristine ng-valid"));
                        passwordField.Clear(); // Xóa trường nếu có giá trị cũ
                        passwordField.SendKeys(password);

                        IWebElement login = driver.FindElement(By.CssSelector("form-group submtit"));
                        login.Click();

                        // Kiểm tra kết quả
                        Thread.Sleep(5000); // Đợi một chút cho các thông báo hiển thị

                        try
                        {
                            IWebElement toastNotification = wait.Until(d => d.FindElement(By.CssSelector("div.Toastify__toast-body")));
                            string toastMessage = toastNotification.Text;
                            Console.WriteLine("Thông báo từ toast: " + toastMessage);

                            // Ghi kết quả vào ô
                            row.CreateCell(2).SetCellValue("Pass"); // Giả sử ghi vào cột 2
                        }
                        catch (WebDriverTimeoutException)
                        {
                            Console.WriteLine("Không có toast xuất hiện.");
                            row.CreateCell(2).SetCellValue("Fail"); // Ghi vào cột 2 nếu không thành công
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