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
    public class AddProduct
    {
        [TestMethod]
        public void TestAddProduct()
        {
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

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
            IWebElement element = driver.FindElement(By.LinkText("Sản phẩm"));
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
                ISheet sheet = workbook.GetSheetAt(7); // Lấy sheet 8

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        IWebElement firstRow = wait.Until(d => d.FindElement(By.XPath("//tbody/tr[1]")));
                        IWebElement detailButton = driver.FindElement(By.XPath(".//button[contains(., 'Chi tiết')]"));
                        detailButton.Click();
                        Thread.Sleep(2000);

                        // Lấy tiêu đề và text hiển thị từ các ô
                        string name = row.GetCell(0)?.ToString(); // Giả sử tiêu đề ở cột 0
                        string weight = row.GetCell(1)?.ToString(); // Giả sử text hiển thị ở cột 1
                        string brand = row.GetCell(2)?.ToString(); // Giả sử text hiển thị ở cột 2
                        string quatity = row.GetCell(3)?.ToString(); // Giả sử text hiển thị ở cột 3
                        string price = row.GetCell(4)?.ToString(); // Giả sử text hiển thị ở cột 4
                        string detail = row.GetCell(5)?.ToString(); // Giả sử text hiển thị ở cột 5

                        IWebElement nameField = driver.FindElement(By.Id("name"));
                        nameField.Clear(); // Xóa trường nếu có giá trị cũ
                        nameField.SendKeys(name);

                        IWebElement weightField = driver.FindElement(By.Id("weight"));
                        weightField.Clear(); // Xóa trường nếu có giá trị cũ
                        weightField.SendKeys(weight);

                        IWebElement brandField = driver.FindElement(By.Name("brand"));
                        brandField.Clear(); // Xóa trường nếu có giá trị cũ
                        brandField.SendKeys(brand);

                        IWebElement priceField = driver.FindElement(By.Name("price"));
                        priceField.Clear(); // Xóa trường nếu có giá trị cũ
                        priceField.SendKeys(price);

                        IWebElement detailField = driver.FindElement(By.Name("detail"));
                        detailField.Clear(); // Xóa trường nếu có giá trị cũ
                        detailField.SendKeys(detail);

                        //IWebElement changeButton = driver.FindElement(By.CssSelector(".label__change__image input[type='file']"));
                        //string imagePath = @"C:\Users\pcgosei\Downloads\img.png";
                        //changeButton.SendKeys(imagePath);
                        //Assert.IsTrue(File.Exists(imagePath));

                        IWebElement dateInput = driver.FindElement(By.CssSelector("input[placeholder='Select date']"));

                        string currentDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        js.ExecuteScript("arguments[0].value='" + currentDateTime + "'; arguments[0].dispatchEvent(new Event('input'));", dateInput);
                        js.ExecuteScript("arguments[0].dispatchEvent(new Event('change'));", dateInput);

                        IWebElement addButton = driver.FindElement(By.XPath("//button[span[text()='Thêm mới']]"));
                        addButton.Click();
                        Thread.Sleep(1000);
                        // Kiểm tra kết quả 
                        Thread.Sleep(1000);

                        if (nameField.GetAttribute("value").Length > 0 && weightField.GetAttribute("value").Length > 0 && brandField.GetAttribute("value").Length > 0 && priceField.GetAttribute("value").Length > 0 && detailField.GetAttribute("value").Length > 0)
                        {
                            if (nameField.GetAttribute("value").Length < 50 && brandField.GetAttribute("value").Length < 50 && detailField.GetAttribute("value").Length < 50)
                            {
                                if (!System.Text.RegularExpressions.Regex.IsMatch(weight, @"^\d+$"))
                                {
                                    Console.WriteLine("Hệ thống vẫn cho phép nhập sai định dạng trường cân nặng");
                                    row.CreateCell(6).SetCellValue("Fail_Hệ thống vẫn cho phép nhập sai định dạng trường cân nặng "); // Giả sử ghi vào cột 2
                                }
                                else if (!System.Text.RegularExpressions.Regex.IsMatch(quatity, @"^\d+$"))
                                {
                                    Console.WriteLine("Hệ thống vẫn cho phép nhập sai định dạng trường số lượng");
                                    row.CreateCell(6).SetCellValue("Fail_Hệ thống vẫn cho phép nhập sai định dạng trường số lượng "); // Giả sử ghi vào cột 2
                                }
                                else if (!System.Text.RegularExpressions.Regex.IsMatch(price, @"^\d+$"))
                                {
                                    Console.WriteLine("Hệ thống vẫn cho phép nhập sai định dạng trường giá gốc");
                                    row.CreateCell(6).SetCellValue("Fail_Hệ thống vẫn cho phép nhập sai định dạng trường giá gốc "); // Giả sử ghi vào cột 2
                                }
                                else
                                {
                                    try
                                    {
                                        IWebElement notificationMessage = driver.FindElement(By.XPath("//div[@role='alert']/div[2]"));
                                        string messageText = notificationMessage.Text;
                                        Console.WriteLine("Notification Message: " + messageText);
                                        row.CreateCell(6).SetCellValue("Pass"); // Giả sử ghi vào cột 2
                                    }
                                    catch (WebDriverTimeoutException)
                                    {
                                        Console.WriteLine("Không có toast xuất hiện.");
                                        row.CreateCell(6).SetCellValue("Fail"); // Giả sử ghi vào cột 2
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Hệ thống vẫn cho nhập quá 50 kí tự vào các trường thông tin");
                                row.CreateCell(6).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 50 kí tự vào các trường thông tin"); // Ghi vào cột 2 
                            }
                        }
                        else
                        {
                            Console.WriteLine("Hệ thống vẫn cho để trống các trường thông tin");
                            row.CreateCell(6).SetCellValue("Fail_Hệ thống vẫn cho phép để trống các trường thông tin"); // Ghi vào cột 2 
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