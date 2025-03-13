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
    public class UpdateProduct
    {
        [TestMethod]
        public void TestUpdateProduct()
        {
            Setup setup = new Setup();
            IWebDriver driver = setup.SetupChromeDriver();
            WebDriverWait wait = setup.CreateWebDriverWait(driver);

            // Đăng nhập
            IWebElement loginButton = wait.Until(d => d.FindElement(By.CssSelector(".MuiBox-root.css-1wenov3")));
            loginButton.Click();

            IWebElement emailField = driver.FindElement(By.CssSelector(".MuiOutlinedInput-input.MuiInputBase-input.css-i46v6x-MuiInputBase-input-MuiOutlinedInput-input"));
            emailField.SendKeys("admin");

            IWebElement passwordField = driver.FindElement(By.Name("password"));
            passwordField.SendKeys("@dmIn12");

            IWebElement login = wait.Until(d => d.FindElement(By.CssSelector(".css-1o3ev6r-MuiButtonBase-root-MuiButton-root")));
            login.Click();
            Thread.Sleep(3000);

            // Chuyển đến trang sản phẩm
            IWebElement product = driver.FindElement(By.XPath("//div[@role='button' and .//div[text()='Sản phẩm']]"));
            product.Click();
            Thread.Sleep(500);

            IWebElement productList = driver.FindElement(By.CssSelector("a[href='/dashboard/app/products/list']"));
            productList.Click();
            Thread.Sleep(15000);

            // Đọc file Excel
            string filePath = @"C:\DATN\datainputandouput.xlsx";
            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại.");
                return;
            }

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(8);

                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        string productName = row.GetCell(0)?.ToString() ?? string.Empty;
                        string productWarranty = row.GetCell(1)?.ToString() ?? string.Empty;
                        string productIntroduce = row.GetCell(2)?.ToString() ?? string.Empty;
                        string productImages = row.GetCell(3)?.ToString() ?? string.Empty;
                        string productTag = row.GetCell(4)?.ToString() ?? string.Empty;

                        IWebElement btnChose = driver.FindElements(By.CssSelector(".MuiButtonBase-root.MuiIconButton-root.MuiIconButton-sizeMedium.css-19fss5u-MuiButtonBase-root-MuiIconButton-root"))[5];
                        btnChose.Click();
                        Thread.Sleep(500);

                        IWebElement btnUpdate = driver.FindElement(By.CssSelector(".MuiButtonBase-root.MuiListItem-root.MuiListItem-gutters.MuiListItem-button.MuiMenuItem-root"));
                        btnUpdate.Click();
                        Thread.Sleep(3000);

                        // Nhập thông tin sản phẩm
                        IWebElement nameField = driver.FindElement(By.Name("name"));
                        nameField.Clear();
                        nameField.SendKeys(Keys.Control + "a");
                        nameField.SendKeys(Keys.Delete);
                        nameField.SendKeys(productName);

                        IWebElement warrantyField = driver.FindElement(By.Name("warrantyPeriod"));
                        warrantyField.Clear();
                        warrantyField.SendKeys(Keys.Control + "a");
                        warrantyField.SendKeys(Keys.Delete);
                        warrantyField.SendKeys(productWarranty);

                        IWebElement introduceField = driver.FindElement(By.CssSelector("div.ql-editor"));
                        introduceField.Clear();
                        introduceField.SendKeys(Keys.Control + "a");
                        introduceField.SendKeys(Keys.Delete);
                        introduceField.SendKeys(productIntroduce);

                        IWebElement imageField = driver.FindElement(By.Name("video"));
                        imageField.SendKeys(Keys.Control + "a");
                        imageField.SendKeys(Keys.Delete);
                        imageField.SendKeys(productImages);
                        imageField.SendKeys(Keys.Enter);

                        // Nhập tag sản phẩm
                        IWebElement tag = driver.FindElements(By.CssSelector(".MuiOutlinedInput-input.MuiInputBase-input.MuiInputBase-inputAdornedEnd.MuiAutocomplete-input.MuiAutocomplete-inputFocused.MuiAutocomplete-input.MuiAutocomplete-inputFocused.css-slontb-MuiInputBase-input-MuiOutlinedInput-input"))[2];
                        tag.SendKeys(Keys.Control + "a");
                        tag.SendKeys(Keys.Delete);
                        tag.SendKeys(productTag);

                        IWebElement addButton = driver.FindElement(By.CssSelector("button.MuiButton-root.MuiButton-contained.MuiButton-containedPrimary.MuiButton-sizeLarge.MuiButton-containedSizeLarge.MuiButton-fullWidth.MuiButtonBase-root.css-1o3ev6r-MuiButtonBase-root-MuiButton-root"));

                        // Kiểm tra điều kiện và ghi kết quả vào file
                        if (productName.Length == 0)
                        {
                            row.CreateCell(5).SetCellValue("Pass");
                            productList.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productName.Length > 25)
                        {
                            row.CreateCell(5).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 25 ký tự");
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productIntroduce.Length == 0)
                        {
                            row.CreateCell(5).SetCellValue("Pass");
                            productList.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productIntroduce.Length > 25)
                        {
                            row.CreateCell(5).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 50 ký tự");
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productWarranty.Length == 0)
                        {
                            row.CreateCell(5).SetCellValue("Pass");
                            productList.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productWarranty.Length >= 25)
                        {
                            row.CreateCell(5).SetCellValue("Fail_Hệ thống vẫn cho phép nhập quá 25 ký tự");
                            addButton.Click();
                            Thread.Sleep(5000);
                        }
                        else if (productImages.Length == 0)
                        {
                            row.CreateCell(5).SetCellValue("Pass");
                            productList.Click();
                            Thread.Sleep(5000);
                        }
                        else
                        {
                            row.CreateCell(5).SetCellValue("Pass");
                            addButton.Click();
                            Thread.Sleep(10000);
                        }
                    }
                }

                // Lưu file Excel
                using (FileStream writeStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(writeStream);
                }
            }

            driver.Quit();
        }


    }
}