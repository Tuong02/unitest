using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.Formula.Functions;
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
    public class UpdateCategory
    {
        [TestMethod]
        public void TestUpdateCategory()
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

            Thread.Sleep(4000);
            var element = driver.FindElement(By.LinkText("Danh mục"));
            element.Click();
            Thread.Sleep(3000);

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
                ISheet sheet = workbook.GetSheetAt(2); // Lấy sheet 3

                for (int i = 1; i <= sheet.LastRowNum; i++)
                { // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        // Lấy tiêu đề và text hiển thị từ các ô
                        string name = row.GetCell(0)?.ToString(); // Giả sử tiêu đề ở cột 0
                        string desc = row.GetCell(1)?.ToString(); // Giả sử text hiển thị ở cột 1

                        IWebElement thirdRow = driver.FindElements(By.CssSelector(".MuiButtonBase-root.MuiIconButton-root.MuiIconButton-sizeMedium.css-19fss5u-MuiButtonBase-root-MuiIconButton-root"))[10];
                        thirdRow.Click();
                        Thread.Sleep(500);

                        // Thực hiện update loại sản phẩm
                        IWebElement webElement = driver.FindElement(By.CssSelector(".MuiButtonBase-root.MuiListItem-root.MuiListItem-gutters.MuiListItem-button.MuiMenuItem-root"));
                        webElement.Click();
                        Thread.Sleep(200);

                        IWebElement categoryInput = driver.FindElement(By.Name("name"));
                        categoryInput.SendKeys(Keys.Control + "a");
                        categoryInput.SendKeys(Keys.Delete);
                        categoryInput.SendKeys(name);

                        IWebElement descInput = driver.FindElement(By.Name("desc"));
                        descInput.SendKeys(Keys.Control + "a");
                        descInput.SendKeys(Keys.Delete);
                        descInput.SendKeys(desc);

                        // Kiểm tra kết quả
                        Thread.Sleep(1000);

                        try
                        {
                            if (name.Length >= 6 && desc.Length > 0 && name.Length <= 50 && desc.Length <= 50)
                            {
                                row.CreateCell(2).SetCellValue("Pass"); // Ghi vào cột 2 nếu thành công
                            }
                            else if (name.Length >= 6 && desc.Length == 0)
                            {
                                row.CreateCell(2).SetCellValue("Fail_Hệ thống vẫn cho phép bỏ trống mô tả");
                            }
                            else if (name.Length > 50 || desc.Length > 50)
                            {
                                row.CreateCell(2).SetCellValue("Fail_Không cho phép vượt quá 50 kí tự");
                            }
                            else
                            {
                                row.CreateCell(2).SetCellValue("Fail_Không đạt yêu cầu");
                            }

                            // Click button Save
                            IWebElement btnSave = driver.FindElement(By.CssSelector(".css-75wtpa-MuiDialogActions-root > :not(:first-of-type)"));
                            btnSave.Click();
                            Thread.Sleep(7000);
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("Lỗi trong quá trình kiểm tra.");
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