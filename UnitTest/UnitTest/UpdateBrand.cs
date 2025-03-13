using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTest
{
    [TestClass]
    public class UpdateBrand
    {
        [TestMethod]
        public void TestUpdateBrand()
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
            Thread.Sleep(15000);

            // click button brand
            IWebElement element = wait.Until(d => d.FindElement(By.LinkText("Thương hiệu")));
            element.Click();
            Thread.Sleep(15000);

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
                ISheet sheet = workbook.GetSheetAt(6); // Lấy sheet 7

                for (int i = 1; i <= sheet.LastRowNum; i++)
                { // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        // Lấy tiêu đề và text hiển thị từ các ô
                        string name = row.GetCell(0)?.ToString(); // Giả sử tiêu đề ở cột 0
                        string desc = row.GetCell(1)?.ToString(); // Giả sử text hiển thị ở cột 1
                        Thread.Sleep(5000);

                        IWebElement chose = wait.Until(d => d.FindElements(By.CssSelector(".MuiButtonBase-root.MuiIconButton-root.MuiIconButton-sizeMedium.css-19fss5u-MuiButtonBase-root-MuiIconButton-root")))[4];
                        chose.Click();
                        //Thread.Sleep(1000);

                        IWebElement deleteButton = driver.FindElement(By.XPath("//span[text()='Chỉnh sửa']"));
                        deleteButton.Click();
                        Thread.Sleep(1000);

                        IWebElement categoryInput = driver.FindElement(By.Name("name"));
                        categoryInput.SendKeys(Keys.Control + "a");
                        categoryInput.SendKeys(Keys.Delete);
                        categoryInput.SendKeys(name);
                        Thread.Sleep(1000);

                        IWebElement descInput = driver.FindElement(By.Name("desc"));
                        descInput.SendKeys(Keys.Control + "a");
                        descInput.SendKeys(Keys.Delete);
                        descInput.SendKeys(desc);
                        Thread.Sleep(1000);

                        try
                        {
                            if (name.Length == 26 && desc.Length >= 1)
                            {
                                row.CreateCell(2).SetCellValue("Pass");
                                IWebElement btnCancel = wait.Until(d => d.FindElement(By.CssSelector(".css-1bacbjs-MuiButtonBase-root-MuiButton-root")));
                                btnCancel.Click();
                                //Thread.Sleep(3000);
                            }
                            else if (name.Length >= 6 && desc.Length == 0)
                            {
                                row.CreateCell(2).SetCellValue("Fail_Hệ thống vẫn cho phép để trống trường mô tả");
                                IWebElement btnSave = wait.Until(d => d.FindElement(By.CssSelector(".css-75wtpa-MuiDialogActions-root > :not(:first-of-type)")));
                                btnSave.Click();
                                //Thread.Sleep(3000);
                            }
                            else if (name.Length >= 6 && desc.Length >= 1)
                            {
                                row.CreateCell(2).SetCellValue("Pass");
                                IWebElement btnSave = wait.Until(d => d.FindElement(By.CssSelector(".css-75wtpa-MuiDialogActions-root > :not(:first-of-type)")));
                                btnSave.Click();
                                //Thread.Sleep(3000);
                            }
                            else
                            {
                                row.CreateCell(2).SetCellValue("Pass");
                                IWebElement btnCancel = wait.Until(d => d.FindElement(By.CssSelector(".css-1bacbjs-MuiButtonBase-root-MuiButton-root")));
                                btnCancel.Click();
                                //Thread.Sleep(3000);
                            }

                        }
                        catch (Exception e)
                        {
                            row.CreateCell(2).SetCellValue("Fail");
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
