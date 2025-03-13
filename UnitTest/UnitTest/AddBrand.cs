using Microsoft.VisualStudio.TestTools.UnitTesting;
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

namespace UnitTest
{
    [TestClass]
    public class AddBrand
    {
        [TestMethod]
        public void TestAddBrand()
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
            Thread.Sleep(10000);

            // click button brand
            IWebElement element = wait.Until(d => d.FindElement(By.LinkText("Thương hiệu")));
            element.Click();
            Thread.Sleep(10000);

            string filePath = @"C:\DATN\datainputandouput.xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(5); // Lấy sheet 6

                for (int i = 1; i <= sheet.LastRowNum; i++) // Bắt đầu từ 1 nếu hàng đầu tiên là tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        
                        // Lấy tên thương hiệu từ cột 0
                        string name = row.GetCell(0)?.ToString();
                        string desc = row.GetCell(1)?.ToString();

                        // click add brand
                        IWebElement addButton = wait.Until(d =>d.FindElement(By.CssSelector(".MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeMedium.MuiButton-containedSizeMedium.MuiButtonBase-root.css-16hkhch-MuiButtonBase-root-MuiButton-root")));
                        addButton.Click();
                        Thread.Sleep(5000);

                        // Nhập tên thương hiệu
                        IWebElement brandName = driver.FindElement(By.XPath("//input[@name='name']"));
                        brandName.SendKeys(Keys.Control + "a");
                        brandName.SendKeys(name);
                        Thread.Sleep(2000);

                        // Nhập mô tả
                        IWebElement descBrand = driver.FindElement(By.XPath("//textarea[@name='desc']"));
                        descBrand.SendKeys(Keys.Control + "a");
                        descBrand.SendKeys(desc);
                        Thread.Sleep(2000);

                        try
                        {
                            if (name.Length == 26 && desc.Length >= 1)
                            {
                                row.CreateCell(2).SetCellValue("Pass");
                                IWebElement btnCancel = driver.FindElement(By.CssSelector(".MuiButton-root.MuiButton-text.MuiButton-textInherit.MuiButton-sizeMedium.MuiButton-textSizeMedium.MuiButton-colorInherit.MuiButtonBase-root.css-1bacbjs-MuiButtonBase-root-MuiButton-root"));
                                btnCancel.Click();
                                Thread.Sleep(2000);
                            }
                            else if (name.Length >= 6 && desc.Length == 0)
                            {
                                row.CreateCell(2).SetCellValue("Fail_Hệ thống vẫn cho phép để trống trường mô tả");
                                IWebElement btnSave = driver.FindElement(By.XPath("//button[span[text()='Thêm']]"));
                                btnSave.Click();
                                Thread.Sleep(2000);
                            }
                            // trên 25 kí tự thì sửa ở đây
                            else if (name.Length >= 6 && desc.Length >= 1)
                            {
                                row.CreateCell(2).SetCellValue("Pass");
                                IWebElement btnSave = driver.FindElement(By.XPath("//button[span[text()='Thêm']]"));
                                btnSave.Click();
                                Thread.Sleep(2000);
                            }
                            else
                            {
                                row.CreateCell(2).SetCellValue("Pass");
                                IWebElement btnCancel = driver.FindElement(By.CssSelector(".MuiButton-root.MuiButton-text.MuiButton-textInherit.MuiButton-sizeMedium.MuiButton-textSizeMedium.MuiButton-colorInherit.MuiButtonBase-root.css-1bacbjs-MuiButtonBase-root-MuiButton-root"));
                                btnCancel.Click();
                                Thread.Sleep(2000);
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
